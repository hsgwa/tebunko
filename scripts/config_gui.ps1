# 画面（WPF）
#
# ［1 インデックス作成］［2 検索］［9 Office 強制終了］の3タブ。画面の定義は config_gui.xaml。
# 変換は office_to_tsv.ps1 をウィンドウを出さずに起動して進み具合を表示し、検索・強制終了は画面内で行う（処理は common.ps1 と共通）。

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

. "$PSScriptRoot\common.ps1"

$ErrorActionPreference = "Stop"

${appTitle}    = "win_grep"
${searchLimit} = 10000
${previewLines} = 3  # 選択行のプレビューに出す前後の行数
${commonPath}  = "$PSScriptRoot\common.ps1"

trap {
    [System.Windows.MessageBox]::Show("予期しないエラーが発生しました。`n$($_.Exception.Message)", ${appTitle}, "OK", "Error") | Out-Null
    exit 1
}

# ---- 多重起動の防止（ツールの配置フォルダごと） ----

$md5 = New-Object System.Security.Cryptography.MD5CryptoServiceProvider
$mutexName = "Local\win_grep_gui_" + [BitConverter]::ToString($md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes(${rootDir}.ToLowerInvariant()))).Replace("-", "")
$createdNew = $false
$mutex = New-Object System.Threading.Mutex($true, $mutexName, [ref]$createdNew)
if (!$createdNew) {
    [System.Windows.MessageBox]::Show("すでに開いています。", ${appTitle}, "OK", "Information") | Out-Null
    exit
}

# ---- 画面で使う型（件数が多くても軽く動くよう C# で定義する） ----

if (!("WinGrep.HitRow" -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace WinGrep
{
    public class Segment
    {
        public string Text { get; set; }
        public bool IsHit { get; set; }
    }

    // 選択行のプレビューの列。見出しの行と各行のセルが同じものを参照するため、
    // 幅を変えると（見出しをドラッグしたとき）列全体の幅が変わる
    public class PreviewColumn : INotifyPropertyChanged
    {
        private double width;

        public string Label { get; set; }
        public double Width
        {
            get { return width; }
            set
            {
                double newWidth = Math.Max(MinWidth, value);
                if (width == newWidth) return;
                width = newWidth;
                PropertyChangedEventHandler handler = PropertyChanged;
                if (handler != null) handler(this, new PropertyChangedEventArgs("Width"));
            }
        }

        public const double MinWidth = 24;  // ドラッグで狭くできる下限

        public event PropertyChangedEventHandler PropertyChanged;
    }

    // 選択行のプレビューのセル1つ（IsSelected はコピーのために選んだ範囲）
    public class PreviewCell : INotifyPropertyChanged
    {
        private bool selected;

        public string Text { get; set; }
        public string ToolTip { get; set; }
        public PreviewColumn Column { get; set; }
        public bool IsHit { get; set; }
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public bool IsSelected
        {
            get { return selected; }
            set
            {
                if (selected == value) return;
                selected = value;
                PropertyChangedEventHandler handler = PropertyChanged;
                if (handler != null) handler(this, new PropertyChangedEventArgs("IsSelected"));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }

    // 選択行のプレビューの1行
    public class PreviewRow
    {
        public string Number { get; set; }
        public bool IsHitRow { get; set; }
        public List<PreviewCell> Cells { get; set; }
    }

    // 選択行のプレビュー。HitOffset・HitWidth は選択行で最初に一致したセルの左端と幅（横スクロール用）。
    // セルをクリックして選んだ範囲（コピー用）も持つ
    public class PreviewTable
    {
        public List<PreviewColumn> Columns { get; set; }
        public List<PreviewRow> Rows { get; set; }
        public double HitOffset { get; set; }
        public double HitWidth { get; set; }

        int anchorRow = -1;
        int anchorColumn = -1;
        int focusRow = -1;
        int focusColumn = -1;

        public bool HasSelection { get { return anchorRow >= 0; } }

        // セルを選ぶ。extend が真なら、選び始めたセルからの四角い範囲にする（Shift＋クリック・ドラッグ）
        public void Select(PreviewCell cell, bool extend)
        {
            if (cell == null) return;
            if (!extend || anchorRow < 0)
            {
                anchorRow = cell.RowIndex;
                anchorColumn = cell.ColumnIndex;
            }
            focusRow = cell.RowIndex;
            focusColumn = cell.ColumnIndex;
            ApplySelection();
        }

        // 1行すべてを選ぶ
        public void SelectRow(PreviewCell cell)
        {
            if (cell == null || Columns.Count == 0) return;
            anchorRow = focusRow = cell.RowIndex;
            anchorColumn = 0;
            focusColumn = Columns.Count - 1;
            ApplySelection();
        }

        public void ClearSelection()
        {
            anchorRow = anchorColumn = focusRow = focusColumn = -1;
            ApplySelection();
        }

        void ApplySelection()
        {
            int rowFrom = Math.Min(anchorRow, focusRow), rowTo = Math.Max(anchorRow, focusRow);
            int columnFrom = Math.Min(anchorColumn, focusColumn), columnTo = Math.Max(anchorColumn, focusColumn);
            foreach (PreviewRow row in Rows)
            {
                foreach (PreviewCell cell in row.Cells)
                {
                    cell.IsSelected = anchorRow >= 0 &&
                                      cell.RowIndex >= rowFrom && cell.RowIndex <= rowTo &&
                                      cell.ColumnIndex >= columnFrom && cell.ColumnIndex <= columnTo;
                }
            }
        }

        // 選んだ範囲の文字列。1 セルならその値のまま、複数ならタブ区切り（Excel に貼り付けたときに
        // 元の位置に並ぶよう、改行・タブ・" を含むセルは " で囲む）
        public string GetSelectionText()
        {
            if (anchorRow < 0) return "";
            int rowFrom = Math.Min(anchorRow, focusRow), rowTo = Math.Max(anchorRow, focusRow);
            int columnFrom = Math.Min(anchorColumn, focusColumn), columnTo = Math.Max(anchorColumn, focusColumn);
            if (rowFrom == rowTo && columnFrom == columnTo)
            {
                return CellText(rowFrom, columnFrom);
            }

            StringBuilder text = new StringBuilder();
            for (int r = rowFrom; r <= rowTo; r++)
            {
                if (r > rowFrom) text.Append("\r\n");
                for (int c = columnFrom; c <= columnTo; c++)
                {
                    if (c > columnFrom) text.Append('\t');
                    text.Append(QuoteForExcel(CellText(r, c)));
                }
            }
            return text.ToString();
        }

        // 選んだセルの数（状況表示用）
        public int SelectedCount
        {
            get
            {
                if (anchorRow < 0) return 0;
                return (Math.Abs(anchorRow - focusRow) + 1) * (Math.Abs(anchorColumn - focusColumn) + 1);
            }
        }

        string CellText(int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || rowIndex >= Rows.Count) return "";
            List<PreviewCell> cells = Rows[rowIndex].Cells;
            return columnIndex >= 0 && columnIndex < cells.Count ? cells[columnIndex].Text : "";
        }

        static string QuoteForExcel(string text)
        {
            if (text.IndexOf('\n') < 0 && text.IndexOf('\t') < 0 && text.IndexOf('"') < 0) return text;
            return "\"" + text.Replace("\"", "\"\"") + "\"";
        }
    }

    // 検索結果の1行（common.ps1 の WinGrep.TsvSearchHit と同じ項目を持ち、toSearchResultLines 等にそのまま渡せる）
    public class HitRow
    {
        public string IndexName { get; set; }
        public string Root { get; set; }
        public string RelPath { get; set; }
        public string RelDir { get; set; }
        public string FileName { get; set; }
        public string Book { get; set; }
        public string Location { get; set; }
        public int LineNumber { get; set; }
        public string Line { get; set; }
        public string DisplayLine { get; set; }
        public List<Segment> Segments { get; set; }
        public bool IsExcel { get; set; }
        public string MatchCell { get; set; }  // 最初に一致した Excel のセル（例: B5）。Excel 以外・見つからなければ空
        public string CellText { get; set; }   // 結果の表の「セル」列（例: B5 / B5 ほか 2）

        private string word;
        private Regex pattern;
        private string filterText;

        // 行の先頭にタブを足してから使い、どのセルも「タブ＋中身」で取る（^ を使うと、先頭の空のセルの後のタブを読み飛ばす）
        static readonly Regex CellRegex = new Regex("\\t(?:\"(?:[^\"]|\"\")*\"[^\\t]*|[^\\t]*)");
        static readonly Regex QuoteRegex = new Regex("^\"((?:[^\"]|\"\")*)\"(.*)$", RegexOptions.Singleline);
        static readonly Regex ExcelRegex = new Regex("\\.xls[a-z]?$", RegexOptions.IgnoreCase);
        const char CellNewLine = (char)0x2028;  // インデックスのTSVでセル内改行の代わりに使う文字（common.ps1 の cellNewLine）
        const int LeadLength = 40;
        const int MaxDisplay = 600;

        public static HitRow Create(string indexName, string root, string relPath, string relDir, string fileName,
                                    string book, string location, int lineNumber, string line, string word, Regex pattern)
        {
            HitRow row = new HitRow();
            row.IndexName = indexName;
            row.Root = root;
            row.RelPath = relPath;
            row.RelDir = relDir ?? "";
            row.FileName = fileName;
            row.Book = book;
            row.Location = location;
            row.LineNumber = lineNumber;
            row.Line = line ?? "";
            row.word = word;
            row.pattern = pattern;
            row.IsExcel = ExcelRegex.IsMatch(book ?? "");
            row.DisplayLine = ToDisplay(row.Line);
            row.Segments = row.BuildSegments();
            row.SetMatchCell();
            row.filterText = row.RelDir + "\t" + book + "\t" + location + "\t" + lineNumber + "\t" + row.CellText + "\t" + row.DisplayLine;
            return row;
        }

        // 一致したセルを数え、MatchCell・CellText を決める
        void SetMatchCell()
        {
            MatchCell = "";
            CellText = "";
            if (!IsExcel) return;
            List<string> cells = SplitCells(Line, true);
            int count = 0;
            for (int i = 0; i < cells.Count; i++)
            {
                if (FindMatches(cells[i], word, pattern).Count == 0) continue;
                if (count == 0) MatchCell = ColumnName(i + 1) + LineNumber;
                count++;
            }
            CellText = count > 1 ? MatchCell + " ほか " + (count - 1) : MatchCell;
        }

        // 絞り込み（全列の部分一致。大文字・小文字を区別しない）
        public bool Contains(string text)
        {
            return filterText.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        // セルの区切り（タブ）とセル内改行を見やすい記号にする
        public static string ToDisplay(string text)
        {
            return (text ?? "").Replace("\t", " │ ").Replace(CellNewLine, '↵');
        }

        public static List<int[]> FindMatches(string text, string word, Regex pattern)
        {
            List<int[]> list = new List<int[]>();
            if (string.IsNullOrEmpty(text)) return list;
            if (pattern != null)
            {
                foreach (Match m in pattern.Matches(text))
                {
                    if (m.Length > 0) list.Add(new int[] { m.Index, m.Length });
                }
            }
            else if (!string.IsNullOrEmpty(word))
            {
                int i = 0;
                while (i < text.Length && (i = text.IndexOf(word, i, StringComparison.CurrentCultureIgnoreCase)) >= 0)
                {
                    list.Add(new int[] { i, word.Length });
                    i += Math.Max(1, word.Length);
                }
            }
            return list;
        }

        // 一致箇所で区切る。最初の一致が見えるよう、その前が長ければ末尾だけを残す
        List<Segment> BuildSegments()
        {
            List<Segment> segments = new List<Segment>();
            int pos = 0;
            int shown = 0;
            foreach (int[] m in FindMatches(Line, word, pattern))
            {
                if (m[0] < pos) continue;
                if (m[0] + m[1] > Line.Length) break;
                string before = Line.Substring(pos, m[0] - pos);
                if (segments.Count == 0 && before.Length > LeadLength)
                {
                    before = "…" + before.Substring(before.Length - LeadLength);
                }
                if (before.Length > 0) segments.Add(new Segment { Text = ToDisplay(before), IsHit = false });
                segments.Add(new Segment { Text = ToDisplay(Line.Substring(m[0], m[1])), IsHit = true });
                shown += before.Length + m[1];
                pos = m[0] + m[1];
                if (shown > MaxDisplay) break;
            }
            if (pos < Line.Length)
            {
                string rest = Line.Substring(pos);
                if (rest.Length > MaxDisplay) rest = rest.Substring(0, MaxDisplay) + "…";
                segments.Add(new Segment { Text = ToDisplay(rest), IsHit = false });
            }
            return segments;
        }

        // TSVの1行をセルに分ける。Excel は " で囲まれたセルを1セルとして囲みを外す。Word・PowerPoint はタブで分けるだけ
        public static List<string> SplitCells(string line, bool isExcel)
        {
            List<string> cells = new List<string>();
            if (!isExcel)
            {
                cells.AddRange((line ?? "").Split('\t'));
                return cells;
            }
            foreach (Match m in CellRegex.Matches("\t" + (line ?? "")))
            {
                string cell = m.Value.Substring(1);
                Match q = QuoteRegex.Match(cell);
                if (q.Success) cell = q.Groups[1].Value.Replace("\"\"", "\"") + q.Groups[2].Value;
                cells.Add(cell);
            }
            return cells;
        }

        const double NumberWidth = 44;     // プレビューの行番号の列の幅（config_gui.xaml と合わせる）
        const double MinCellWidth = 48;
        const double MaxCellWidth = 260;
        const double MaxParagraphWidth = 640;

        // 選択行のプレビュー（前後の行をセルに分けた、シートのような表）を作る。
        // numbers・lines は TSV の行番号と行（common.ps1 の readTsvContext）。読めなかった場合は選択行だけにする
        public PreviewTable BuildPreview(int[] numbers, string[] lines)
        {
            if (numbers == null || lines == null || numbers.Length == 0 || numbers.Length != lines.Length)
            {
                numbers = new int[] { LineNumber };
                lines = new string[] { Line };
            }

            List<List<string>> rowCells = new List<List<string>>();
            int columnCount = 0;
            foreach (string line in lines)
            {
                List<string> cells = SplitCells(line, IsExcel);
                rowCells.Add(cells);
                columnCount = Math.Max(columnCount, cells.Count);
            }

            // 列の幅は、見出しと各行のセルの文字数（セル内改行のある行は最も長い行）から決める。
            // Word・PowerPoint は段落が1列になるため広くする。ここで決めるのは初めの幅で、見出しのドラッグで変えられる
            double maxWidth = IsExcel ? MaxCellWidth : MaxParagraphWidth;
            PreviewTable table = new PreviewTable
            {
                Columns = new List<PreviewColumn>(), Rows = new List<PreviewRow>(), HitOffset = -1, HitWidth = 0
            };
            for (int c = 0; c < columnCount; c++)
            {
                double width = TextWidth(ColumnLabel(c + 1));
                foreach (List<string> cells in rowCells)
                {
                    if (c < cells.Count) width = Math.Max(width, CellWidth(cells[c]));
                }
                table.Columns.Add(new PreviewColumn
                {
                    Label = ColumnLabel(c + 1),
                    Width = Math.Min(maxWidth, Math.Max(MinCellWidth, width))
                });
            }

            for (int r = 0; r < rowCells.Count; r++)
            {
                bool isHitRow = numbers[r] == LineNumber;
                PreviewRow row = new PreviewRow { Number = numbers[r].ToString(), IsHitRow = isHitRow, Cells = new List<PreviewCell>() };
                double left = NumberWidth;
                for (int c = 0; c < columnCount; c++)
                {
                    string cell = c < rowCells[r].Count ? rowCells[r][c] : "";
                    bool isHit = FindMatches(cell, word, pattern).Count > 0;
                    // セル内改行は改行のまま表示する（折り返して複数行で見えるようにする）
                    string text = cell.Replace(CellNewLine, '\n');
                    row.Cells.Add(new PreviewCell
                    {
                        Text = text,
                        ToolTip = text.Length > 0 ? text : null,
                        Column = table.Columns[c],
                        IsHit = isHit,
                        RowIndex = r,
                        ColumnIndex = c
                    });
                    if (isHitRow && isHit && table.HitOffset < 0)
                    {
                        table.HitOffset = left;
                        table.HitWidth = table.Columns[c].Width;
                    }
                    left += table.Columns[c].Width;
                }
                table.Rows.Add(row);
            }
            if (table.HitOffset < 0) table.HitOffset = 0;
            return table;
        }

        // セルの幅の目安。セル内改行を含むセルは、最も長い行に合わせる
        static double CellWidth(string cell)
        {
            double width = 0;
            foreach (string line in (cell ?? "").Split(CellNewLine, '\n'))
            {
                width = Math.Max(width, TextWidth(line));
            }
            return width;
        }

        // 列見出し。Excel は A, B, C…、Word・PowerPoint はタブで分けた順の 1, 2, 3…
        string ColumnLabel(int number)
        {
            return IsExcel ? ColumnName(number) : number.ToString();
        }

        // 12px の文字で表示したときの幅の目安（半角 7px、全角 12px、左右の余白 14px）
        static double TextWidth(string text)
        {
            double width = 14;
            foreach (char ch in text ?? "")
            {
                width += ch < 0x0100 || (ch >= 0xFF61 && ch <= 0xFF9F) ? 7 : 12;
            }
            return width;
        }

        public static string ColumnName(int number)
        {
            string name = "";
            while (number > 0)
            {
                number--;
                name = (char)('A' + number % 26) + name;
                number /= 26;
            }
            return name;
        }
    }

    // ［9 Office 強制終了］の1行
    public class ProcRow
    {
        public int Id { get; set; }
        public string AppName { get; set; }
        public bool Background { get; set; }
        public string StateText { get; set; }
        public string StartText { get; set; }
        public string MemoryText { get; set; }
        public string TitleText { get; set; }
    }

    // ［1 インデックス作成］の変換に失敗したファイル1件
    public class FailRow
    {
        public string RelPath { get; set; }
        public string Reason { get; set; }
        public string ConvertedText { get; set; }
        public string SourcePath { get; set; }
    }

    // ［1 インデックス作成］の変換対象フォルダ1件（チェックの変更を通知する）
    public class FolderItem : INotifyPropertyChanged
    {
        private bool enabled;
        private string statusText;
        private object statusBrush;
        private string path;

        // インデックス名（work\index 直下のフォルダ名）。フォルダの置き場所（Path）とは分けて持つ
        private string name;
        public string Name
        {
            get { return name; }
            set { if (name != value) { name = value; OnChanged("Name"); OnChanged("IndexLabel"); } }
        }
        // 一覧に出すインデックス名（まだ決まっていなければ空）
        public string IndexLabel
        {
            get { return string.IsNullOrEmpty(name) ? "" : "[" + name + "]"; }
        }
        public string Path
        {
            get { return path; }
            set { if (path != value) { path = value; OnChanged("Path"); } }
        }
        public bool Enabled
        {
            get { return enabled; }
            set { if (enabled != value) { enabled = value; OnChanged("Enabled"); } }
        }
        public string StatusText
        {
            get { return statusText; }
            set { statusText = value; OnChanged("StatusText"); }
        }
        public object StatusBrush
        {
            get { return statusBrush; }
            set { statusBrush = value; OnChanged("StatusBrush"); }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        void OnChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(name));
        }
    }

    public static class Native
    {
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
    }

    // エクスプローラー形式のフォルダ選択（Windows 標準の IFileOpenDialog をフォルダ選択モードで使う）。
    // アドレスバーや「フォルダー」欄にパスを貼り付けて選べる。.NET の FolderBrowserDialog（ツリー形式）ではパスを入力できないため
    public static class FolderPicker
    {
        const uint FOS_NOCHANGEDIR = 0x8;
        const uint FOS_PICKFOLDERS = 0x20;
        const uint FOS_FORCEFILESYSTEM = 0x40;
        const uint FOS_PATHMUSTEXIST = 0x800;
        const uint SIGDN_FILESYSPATH = 0x80058000;
        const int ERROR_CANCELLED = unchecked((int)0x800704C7);

        [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
        class FileOpenDialogCoClass { }

        [ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IFileOpenDialog
        {
            [PreserveSig] int Show(IntPtr parent);
            void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
            void SetFileTypeIndex(uint iFileType);
            void GetFileTypeIndex(out uint piFileType);
            void Advise(IntPtr pfde, out uint pdwCookie);
            void Unadvise(uint dwCookie);
            void SetOptions(uint fos);
            void GetOptions(out uint pfos);
            void SetDefaultFolder(IShellItem psi);
            void SetFolder(IShellItem psi);
            void GetFolder(out IShellItem ppsi);
            void GetCurrentSelection(out IShellItem ppsi);
            void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
            void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
            void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
            void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
            void GetResult(out IShellItem ppsi);
            void AddPlace(IShellItem psi, int fdap);
            void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
            void Close(int hr);
            void SetClientGuid(ref Guid guid);
            void ClearClientData();
            void SetFilter(IntPtr pFilter);
            void GetResults(out IntPtr ppenum);
            void GetSelectedItems(out IntPtr ppsai);
        }

        [ComImport, Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IShellItem
        {
            void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
            void GetParent(out IShellItem ppsi);
            void GetDisplayName(uint sigdnName, out IntPtr ppszName);
            void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
            void Compare(IShellItem psi, uint hint, out int piOrder);
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        static extern void SHCreateItemFromParsingName(string pszPath, IntPtr pbc, [In] ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out IShellItem ppv);

        // 選んだフォルダのパスを返す。キャンセルなら null
        public static string Show(IntPtr owner, string title, string initialFolder)
        {
            IFileOpenDialog dialog = (IFileOpenDialog)new FileOpenDialogCoClass();
            try
            {
                uint options;
                dialog.GetOptions(out options);
                dialog.SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_NOCHANGEDIR);
                if (!string.IsNullOrEmpty(title)) dialog.SetTitle(title);
                if (!string.IsNullOrEmpty(initialFolder) && System.IO.Directory.Exists(initialFolder))
                {
                    Guid shellItemId = typeof(IShellItem).GUID;
                    IShellItem folder;
                    SHCreateItemFromParsingName(initialFolder, IntPtr.Zero, ref shellItemId, out folder);
                    dialog.SetFolder(folder);
                }

                int hr = dialog.Show(owner);
                if (hr == ERROR_CANCELLED) return null;
                if (hr != 0) Marshal.ThrowExceptionForHR(hr);

                IShellItem result;
                dialog.GetResult(out result);
                IntPtr namePtr;
                result.GetDisplayName(SIGDN_FILESYSPATH, out namePtr);
                try
                {
                    return Marshal.PtrToStringUni(namePtr);
                }
                finally
                {
                    Marshal.FreeCoTaskMem(namePtr);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(dialog);
            }
        }
    }

    // 検索対象の1件（common.ps1 の getIndexTsvFiles に渡す）
    public class SearchTarget
    {
        public string Root { get; set; }     // 検索対象インデックスのフォルダ
        public string RelPath { get; set; }  // Root の中のフォルダ（空は Root 自身）
        public bool Recurse { get; set; }    // false はフォルダ直下のファイルだけ
    }

    // 検索対象から外したフォルダ（common.ps1 の readSearchExcludes / writeSearchExcludes と同じ項目）
    public class SearchExclude
    {
        public string Path { get; set; }
        public bool Subfolders { get; set; }  // false はフォルダ直下のファイルだけを外す
    }

    // 検索対象インデックスのツリーの1項目。フォルダ、またはフォルダ直下のファイルのまとまり（IsFiles。
    // サブフォルダとファイルの両方があるフォルダだけに付け、サブフォルダの一部だけを選んだときに直下のファイルを含めるかを表す）。
    // チェックは true / false / null（一部）の3状態で、子のあるフォルダの状態は子の状態から決まる。
    // 子のフォルダは展開したときに読み込む（読み込むまでは、下のフォルダもすべて同じ状態とみなす）
    public class IndexNode : INotifyPropertyChanged
    {
        public string Name { get; private set; }
        public string Root { get; private set; }     // 検索対象インデックスのフォルダ（フルパス）
        public string RelPath { get; private set; }  // Root からの相対パス（Root 自身と、Root 直下のファイルは空）
        public bool IsFiles { get; private set; }
        public bool IsPlaceholder { get; private set; }  // 展開するまで置いておく仮の子
        public bool Exists { get; private set; }
        public string SourcePath { get; private set; }   // 元のフォルダ（分からなければ null）
        public string ToolTip { get; private set; }
        public IndexNode Parent { get; private set; }
        public System.Collections.ObjectModel.ObservableCollection<IndexNode> Children { get; private set; }

        IDictionary<string, string> sourceFolders;  // ルートだけ持つ: インデックス名 → 変換対象フォルダ
        bool? isChecked = true;
        bool isExpanded;
        bool loaded;

        public event PropertyChangedEventHandler PropertyChanged;
        void OnChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(name));
        }

        IndexNode(IndexNode parent, string name, string root, string relPath, bool isFiles)
        {
            Parent = parent;
            Name = name;
            Root = root;
            RelPath = relPath;
            IsFiles = isFiles;
            Exists = true;
            Children = new System.Collections.ObjectModel.ObservableCollection<IndexNode>();
            if (parent != null) isChecked = parent.isChecked != false;
        }

        // 検索対象インデックスのフォルダ（ツリーの一番上）を作る。sourceFolders は 元のフォルダ.txt・変換一覧の記録
        public static IndexNode CreateRoot(string root, string name, string sourcePath, IDictionary<string, string> sourceFolders)
        {
            IndexNode node = new IndexNode(null, name, root, "", false);
            node.sourceFolders = sourceFolders;
            node.SourcePath = sourcePath;
            node.Exists = System.IO.Directory.Exists(LongPath(root));
            node.ToolTip = node.Exists ? root : root + "（フォルダが見つかりません。検索時はスキップします）";
            if (sourcePath != null) node.ToolTip += "\n元のフォルダ：" + sourcePath;
            if (node.Exists && HasSubfolders(root)) node.Children.Add(NewPlaceholder(node));
            return node;
        }

        public string FullPath
        {
            get { return RelPath == "" ? Root : Root.TrimEnd('\\') + "\\" + RelPath; }
        }

        public bool? IsChecked
        {
            get { return isChecked; }
            set { SetChecked(value != false); }
        }

        public bool IsExpanded
        {
            get { return isExpanded; }
            set
            {
                if (value) LoadChildren();
                if (isExpanded == value) return;
                isExpanded = value;
                OnChanged("IsExpanded");
            }
        }

        // チェックを付ける・外す（下のフォルダも同じにし、上のフォルダの状態を決め直す）
        public void SetChecked(bool value)
        {
            SetTree(value);
            if (Parent != null) Parent.UpdateFromChildren();
        }

        public void Toggle()
        {
            SetChecked(isChecked != true);
        }

        void SetTree(bool value)
        {
            SetState(value);
            foreach (IndexNode child in Children)
            {
                if (!child.IsPlaceholder) child.SetTree(value);
            }
        }

        void SetState(bool? value)
        {
            if (isChecked == value) return;
            isChecked = value;
            OnChanged("IsChecked");
        }

        void UpdateFromChildren()
        {
            bool any = false;
            bool? state = null;
            foreach (IndexNode child in Children)
            {
                if (child.IsPlaceholder) continue;
                if (!any)
                {
                    state = child.isChecked;
                    any = true;
                }
                else if (state != child.isChecked)
                {
                    state = null;
                    break;
                }
            }
            if (!any) return;
            SetState(state);
            if (Parent != null) Parent.UpdateFromChildren();
        }

        // 子のフォルダを読み込む（1回だけ）。子は今の状態（チェックあり・なし）を引き継ぐ
        public void LoadChildren()
        {
            if (loaded || IsFiles || IsPlaceholder) return;
            loaded = true;
            Children.Clear();
            if (!Exists) return;

            string dir = FullPath;
            List<string> names = new List<string>();
            try
            {
                foreach (string sub in System.IO.Directory.EnumerateDirectories(LongPath(dir)))
                {
                    string name = System.IO.Path.GetFileName(sub);
                    // 元のファイル名のフォルダはツリーに出さない（中のTSVは、このフォルダ直下のファイルとして扱う）
                    if (IsBookDir(name)) continue;
                    names.Add(name);
                }
            }
            catch (Exception)
            {
            }
            names.Sort(StringComparer.CurrentCultureIgnoreCase);

            if (names.Count > 0 && HasFiles(dir))
            {
                IndexNode files = new IndexNode(this, "（このフォルダ直下のファイル）", Root, RelPath, true);
                files.ToolTip = "サブフォルダを除く、" + (SourcePath ?? dir) + " の直下のファイル";
                Children.Add(files);
            }
            foreach (string name in names)
            {
                IndexNode child = new IndexNode(this, name, Root, RelPath == "" ? name : RelPath + "\\" + name, false);
                if (Parent == null && sourceFolders != null && sourceFolders.ContainsKey(name))
                {
                    child.SourcePath = sourceFolders[name];
                }
                else if (SourcePath != null)
                {
                    child.SourcePath = SourcePath.TrimEnd('\\') + "\\" + name;
                }
                child.ToolTip = child.SourcePath != null ? "元のフォルダ：" + child.SourcePath : child.FullPath;
                if (HasSubfolders(child.FullPath)) child.Children.Add(NewPlaceholder(child));
                Children.Add(child);
            }
        }

        // path（フルパス）のフォルダの項目を返す（途中のフォルダは読み込む）。無ければ null
        public IndexNode Find(string path)
        {
            if (IsFiles || IsPlaceholder) return null;
            string full = FullPath.TrimEnd('\\');
            path = path.TrimEnd('\\');
            if (string.Equals(path, full, StringComparison.OrdinalIgnoreCase)) return this;
            if (!path.StartsWith(full + "\\", StringComparison.OrdinalIgnoreCase)) return null;
            LoadChildren();
            foreach (IndexNode child in Children)
            {
                IndexNode found = child.Find(path);
                if (found != null) return found;
            }
            return null;
        }

        // 保存したチェックなしのフォルダを戻す（subfolders が false ならフォルダ直下のファイルだけ）
        public void ApplyExclude(string path, bool subfolders)
        {
            IndexNode node = Find(path);
            if (node == null) return;
            if (!subfolders)
            {
                node.LoadChildren();
                foreach (IndexNode child in node.Children)
                {
                    if (child.IsFiles)
                    {
                        child.SetChecked(false);
                        return;
                    }
                }
                // サブフォルダだけになっていれば直下のファイルは無い。フォルダだけになっていればフォルダごと外す
                if (node.Children.Count > 0) return;
            }
            node.SetChecked(false);
        }

        // 検索する範囲（チェックありの一番上のフォルダ・直下のファイル）を targets に加える
        public void AddTargets(List<SearchTarget> targets)
        {
            if (IsPlaceholder || isChecked == false) return;
            if (isChecked == true)
            {
                targets.Add(new SearchTarget { Root = Root, RelPath = RelPath, Recurse = !IsFiles });
                return;
            }
            foreach (IndexNode child in Children) child.AddTargets(targets);
        }

        // 保存する、チェックなしの一番上のフォルダ・直下のファイルを excludes に加える
        public void AddExcludes(List<SearchExclude> excludes)
        {
            if (IsPlaceholder || isChecked == true) return;
            if (isChecked == false)
            {
                excludes.Add(new SearchExclude { Path = FullPath, Subfolders = !IsFiles });
                return;
            }
            foreach (IndexNode child in Children) child.AddExcludes(excludes);
        }

        // 展開しているフォルダのパスを paths に加える（読み込み直した後に展開の状態を戻すため）
        public void AddExpanded(List<string> paths)
        {
            if (IsFiles || IsPlaceholder) return;
            if (isExpanded) paths.Add(FullPath);
            foreach (IndexNode child in Children) child.AddExpanded(paths);
        }

        static IndexNode NewPlaceholder(IndexNode parent)
        {
            IndexNode node = new IndexNode(parent, "読み込み中…", parent.Root, parent.RelPath, false);
            node.IsPlaceholder = true;
            return node;
        }

        // 260文字を超えるパスも扱えるよう \\?\ を付ける（common.ps1 の toLongPath と同じ）
        static string LongPath(string path)
        {
            if (path.StartsWith("\\\\?\\")) return path;
            if (path.EndsWith(":")) path += "\\";
            if (path.StartsWith("\\\\")) return "\\\\?\\UNC\\" + path.Substring(2);
            return "\\\\?\\" + path;
        }

        // インデックスの「元のファイル名のフォルダ」（<ファイル名.xlsx>\<場所>.tsv のフォルダ）かどうか。
        // 元のファイルはツリーに出さず、そのTSVは親フォルダ直下のファイルとして扱う
        public static bool IsBookDir(string name)
        {
            string ext = System.IO.Path.GetExtension(name).ToLowerInvariant();
            if (ext.Length < 4 || ext.Length > 5) return false;
            return ext.StartsWith(".xls") || ext.StartsWith(".doc") || ext.StartsWith(".ppt");
        }

        static bool HasSubfolders(string dir)
        {
            try
            {
                foreach (string sub in System.IO.Directory.EnumerateDirectories(LongPath(dir)))
                {
                    if (!IsBookDir(System.IO.Path.GetFileName(sub))) return true;
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        static bool HasFiles(string dir)
        {
            try
            {
                using (IEnumerator<string> e = System.IO.Directory.EnumerateFiles(LongPath(dir), "*.tsv").GetEnumerator())
                {
                    if (e.MoveNext()) return true;
                }
                // 今の形式では、TSVは元のファイル名のフォルダの中にある
                foreach (string sub in System.IO.Directory.EnumerateDirectories(LongPath(dir)))
                {
                    if (!IsBookDir(System.IO.Path.GetFileName(sub))) continue;
                    using (IEnumerator<string> e = System.IO.Directory.EnumerateFiles(sub, "*.tsv").GetEnumerator())
                    {
                        if (e.MoveNext()) return true;
                    }
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
'@
}

# ---- 画面の読み込み ----

function loadWindow {
    param (
        [string]$path
    )

    [xml]$xaml = [System.IO.File]::ReadAllText($path)
    return [System.Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
}

$window = loadWindow "$PSScriptRoot\config_gui.xaml"
$ui = @{}
foreach ($name in @(
        "Tabs", "IndexTab", "SearchTab", "KillTab", "IndexTabHeader", "KillTabHeader", "StatusText", "CloseButton",
        "NewFolderBox", "NewFolderPlaceholder", "AddFolderButton", "BrowseFolderButton", "TargetList", "TargetPlaceholder", "RemoveFolderButton",
        "ChangeFolderPathButton",
        "IndexSummaryText", "ConversionStateText", "ConvertButton", "ConvertHint",
        "FailedPanel", "FailedHeading", "FailedGrid",
        "ConvertProgressPanel", "ConvertProgressText", "ConvertProgressEta", "ConvertProgress", "ConvertProgressDetail", "ConvertStopButton", "ConvertLogButton",
        "WordBox", "SearchButton", "RegexCheck", "CaseCheck", "FileFilterBox", "FileFilterPlaceholder", "WordNotice", "SearchTargetText", "ChangeIndexButton", "GoIndexTabButton",
        "IndexTree", "CheckAllIndexButton", "UncheckAllIndexButton",
        "SummaryText", "SearchProgress", "FilterBox", "FilterPlaceholder", "ResultGrid", "IndexColumn",
        "MenuOpen", "MenuOpenReadOnly", "MenuOpenNew", "MenuOpenFolder", "MenuCopy", "MenuCopyPath", "DetailPanel", "DetailTitle", "OpenButton", "OpenModeCombo", "OpenFolderButton", "PreviewScroll", "PreviewHeader", "PreviewRows", "MenuPreviewCopy", "MenuPreviewCopyRow", "ExportButton",
        "ProcessGrid", "ProcessSummaryText", "RefreshProcessButton", "KillAllButton", "KillSelectedButton", "KillBackgroundButton")) {
    $ui[$name] = $window.FindName($name)
}
$taskbar = $window.TaskbarItemInfo

function toBrush {
    param (
        [string]$hex
    )

    return New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString($hex))
}

${okBrush}   = toBrush "#2E8B57"
${warnBrush} = toBrush "#B45309"
${ngBrush}   = toBrush "#DC2626"

# ---- 共通の部品 ----

function setStatus {
    param (
        [string]$text
    )

    $ui.StatusText.Text = $text
    $ui.StatusText.ToolTip = $text
}

function showMessage {
    param (
        [string]$message,
        [string]$buttons = "OK",
        [string]$icon = "Information",
        [string]$default = "None",
        [System.Windows.Window]$owner = $window  # ダイアログを開いているときは、そのダイアログを親にする
    )

    return [System.Windows.MessageBox]::Show($owner, $message, ${appTitle}, $buttons, $icon, $default)
}

function safe {
    # イベント処理で例外が起きても画面を落とさず、内容を表示する
    param (
        [scriptblock]$block
    )

    try {
        & $block
    } catch {
        setStatus "エラーが発生しました：$($_.Exception.Message)"
        showMessage "エラーが発生しました。`n$($_.Exception.Message)" "OK" "Error" | Out-Null
    }
}

function newTimer {
    param (
        [int]$milliseconds,
        [scriptblock]$onTick
    )

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds($milliseconds)
    $timer.Add_Tick($onTick)
    return $timer
}

function selectFolder {
    # フォルダ選択ダイアログ。アドレスバー・「フォルダー」欄にパスを貼り付けて選べるエクスプローラー形式を使い、
    # 使えない環境では .NET 標準のツリー形式にする。キャンセルなら $null
    param (
        [string]$description,
        [string]$initialPath,
        [System.Windows.Window]$owner = $window
    )

    try {
        $handle = (New-Object System.Windows.Interop.WindowInteropHelper($owner)).Handle
        return [WinGrep.FolderPicker]::Show($handle, $description, $initialPath)
    } catch {
    }

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $description
    $dialog.ShowNewFolderButton = $false
    if ($initialPath -and (Test-Path -LiteralPath $initialPath -PathType Container)) {
        $dialog.SelectedPath = $initialPath
    }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.SelectedPath
    }
    return $null
}

function getDroppedFolders {
    param (
        [System.Windows.DragEventArgs]$e
    )

    if (!$e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        return @()
    }
    return @($e.Data.GetData([System.Windows.DataFormats]::FileDrop) | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
}

function onFolderDragOver {
    param ($sender, [System.Windows.DragEventArgs]$e)

    $e.Effects = if ((getDroppedFolders $e).Count -gt 0) { "Copy" } else { "None" }
    $e.Handled = $true
}

function getTargetsKey {
    # 変換対象フォルダの一覧（@{ Path; Enabled } の配列）を比べるための文字列
    param (
        [object[]]$folders
    )

    return (@($folders | Where-Object { $_ } | ForEach-Object { "$($_.Enabled)`t$($_.Path)" }) -join "`n")
}

function formatTime {
    # 当日なら HH:mm、それ以前は M/d HH:mm
    param (
        $time
    )

    if ($null -eq $time) {
        return ""
    }
    if ($time.Date -eq (Get-Date).Date) {
        return $time.ToString("H:mm")
    }
    return $time.ToString("M/d H:mm")
}

function readTextShared {
    # 変換側が書き込み中でも妨げないよう、共有を許して読む
    param (
        [string]$path
    )

    if (!(Test-Path -LiteralPath $path)) {
        return ""
    }
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = New-Object System.IO.FileStream($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    $reader = New-Object System.IO.StreamReader($stream, ${utf8Bom})
    try {
        return $reader.ReadToEnd()
    } finally {
        $reader.Dispose()
    }
}

# ---- 別スレッドの処理（インデックスの件数など） ----

$script:jobs = New-Object System.Collections.ArrayList

function startJob {
    # scriptBlock を別スレッドで実行し、終わったら画面のスレッドで onDone { param($output, $errorText) } を呼ぶ
    param (
        [scriptblock]$scriptBlock,
        [object[]]$arguments,
        [scriptblock]$onDone
    )

    $ps = [powershell]::Create()
    [void]$ps.AddScript($scriptBlock.ToString())
    foreach ($argument in $arguments) {
        [void]$ps.AddArgument($argument)
    }
    [void]$script:jobs.Add(@{ PS = $ps; Handle = $ps.BeginInvoke(); OnDone = $onDone })
    $script:jobTimer.Start()
}

$script:jobTimer = newTimer 200 {
    safe {
        foreach ($job in @($script:jobs.ToArray())) {
            if (!$job.Handle.IsCompleted) {
                continue
            }
            $script:jobs.Remove($job)
            $output = $null
            $errorText = $null
            try {
                $output = $job.PS.EndInvoke($job.Handle)
                if ($job.PS.Streams.Error.Count -gt 0) {
                    $errorText = $job.PS.Streams.Error[0].ToString()
                }
            } catch {
                $errorText = $_.Exception.Message
            } finally {
                $job.PS.Dispose()
            }
            & $job.OnDone $output $errorText
        }
        if ($script:jobs.Count -eq 0) {
            $script:jobTimer.Stop()
        }
    }
}

# ============================================================================
# ［1 インデックス作成］
# ============================================================================

$script:targetItems = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.TargetList.ItemsSource = $script:targetItems
$script:loadingTargets = $false
$script:savedTargets = $null  # 最後に読み込み・保存した変換対象フォルダ（getTargetsKey）。ほかでの変更の検出に使う
$script:convertProcess = $null
$script:convertStart = $null
$script:convertFailed = 0  # 変換中に一覧へ反映済みの失敗件数
$script:conversionState = $null
$script:indexSummary = $null

function isConverting {
    return ($null -ne $script:convertProcess) -and !$script:convertProcess.HasExited
}

function updateFolderItemStatus {
    param (
        $item
    )

    if (Test-Path -LiteralPath $item.Path -PathType Container) {
        $item.StatusText = "✓ フォルダがあります"
        $item.StatusBrush = ${okBrush}
    } else {
        $item.StatusText = "✗ フォルダが見つかりません"
        $item.StatusBrush = ${ngBrush}
    }
}

function newFolderItem {
    param (
        [string]$path,
        [bool]$enabled,
        [string]$name = ""
    )

    $item = New-Object WinGrep.FolderItem
    $item.Name = $name
    $item.Path = $path
    $item.Enabled = $enabled
    updateFolderItemStatus $item
    $item.Add_PropertyChanged({
        param ($sender, $e)
        if ($e.PropertyName -eq "Enabled" -and !$script:loadingTargets) {
            safe {
                saveTargets
                updateConvertButton
            }
        }
    })
    return $item
}

function loadTargets {
    $script:loadingTargets = $true
    try {
        $script:targetItems.Clear()
        $folders = @(getTargetFolders)
        foreach ($folder in $folders) {
            $script:targetItems.Add((newFolderItem $folder.Path $folder.Enabled $folder.Name))
        }
        $script:savedTargets = getTargetsKey $folders
    } finally {
        $script:loadingTargets = $false
    }
    updateTargetView
}

function saveTargets {
    writeTargetFolders @($script:targetItems | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Path = $_.Path; Enabled = $_.Enabled } })
    $script:savedTargets = getTargetsKey @(getTargetFolders)
    setStatus "変換対象フォルダを保存しました（$(Get-Date -Format 'H:mm')）"
}

function changeTargetFolderPath {
    # 選んだ変換対象フォルダの「場所」だけを変える（インデックス名はそのまま）。
    # フォルダを別のドライブ・共有フォルダへ移したときに、インデックスを作り直さずに済ませるための操作
    $item = $ui.TargetList.SelectedItem
    if ($null -eq $item) {
        return
    }

    $initial = if (Test-Path -LiteralPath $item.Path -PathType Container) { $item.Path } else { getExistingFolder $item.Path }
    $picked = selectFolder "「$($item.Path)」を移した先のフォルダを選んでください" $initial
    if (!$picked) {
        return
    }
    $picked = normalizeFolderPath $picked
    if ($picked -eq $item.Path) {
        return
    }
    foreach ($other in $script:targetItems) {
        if ($other -ne $item -and $other.Path -eq $picked) {
            showMessage "「${picked}」は既に変換対象フォルダにあります。" "OK" "Warning" | Out-Null
            return
        }
    }

    $indexName = if ($item.Name) { "インデックス [$($item.Name)]" } else { "このフォルダのインデックス" }
    $message = "変換対象フォルダの場所を変えます。`n`n変更前：$($item.Path)`n変更後：${picked}`n`n" +
               "${indexName} はそのまま使います（作り直しません）。次の変換では、更新されたファイルだけを変換します。`n`n変更しますか？"
    if ((showMessage $message "YesNo" "Question" "Yes") -ne "Yes") {
        return
    }

    $item.Path = $picked
    updateFolderItemStatus $item
    saveTargets
    $script:sourceFolderMaps = @{}
    refreshConversionState
    setStatus "変換対象フォルダの場所を変えました（${picked}）"
}

function updateTargetView {
    $ui.TargetPlaceholder.Visibility = if ($script:targetItems.Count -eq 0) { "Visible" } else { "Collapsed" }
    $ui.RemoveFolderButton.IsEnabled = $null -ne $ui.TargetList.SelectedItem
    $ui.ChangeFolderPathButton.IsEnabled = $null -ne $ui.TargetList.SelectedItem
    updateConvertButton
}

function addTargetFolder {
    param (
        [string]$path
    )

    $path = normalizeFolderPath $path
    if ($path -eq "") {
        return
    }
    foreach ($item in $script:targetItems) {
        # 書き方が違うだけで同じフォルダ（ネットワークドライブと UNC パスなど）も、すでに登録されているとみなす
        if (testSameFolder $item.Path $path) {
            $ui.TargetList.SelectedItem = $item
            setStatus "すでに登録されています：$($item.Path)"
            return
        }
    }

    $item = newFolderItem $path $true
    $script:targetItems.Add($item)
    $ui.TargetList.SelectedItem = $item
    saveTargets
    if (!(Test-Path -LiteralPath $path -PathType Container)) {
        setStatus "追加しましたが、フォルダが見つかりません：${path}"
    }
    updateTargetView
}

function removeTargetFolder {
    $item = $ui.TargetList.SelectedItem
    if ($null -eq $item) {
        return
    }
    $answer = showMessage ("「$($item.Path)」を変換対象から削除します。`n" +
        "このフォルダのインデックス（検索用に変換したデータ）も、次に変換したときに削除されます。`n" +
        "一時的に変換しないだけなら、削除せずにチェックを外してください。`n`n削除しますか？") "YesNo" "Question" "No"
    if ($answer -ne "Yes") {
        return
    }
    $script:targetItems.Remove($item)
    saveTargets
    updateTargetView
}

function updateConvertButton {
    $ready = $false
    foreach ($item in $script:targetItems) {
        if ($item.Enabled -and (Test-Path -LiteralPath $item.Path -PathType Container)) {
            $ready = $true
            break
        }
    }

    $state = $script:conversionState
    if (isConverting) {
        $ui.ConvertButton.Content = "変換中…"
        $ui.ConvertButton.IsEnabled = $false
    } else {
        $ui.ConvertButton.Content = if ($state -and $state.Pending -gt 0) { "続きから再開（残り $($state.Pending) 件）" } else { "変換を開始" }
        $ui.ConvertButton.IsEnabled = $ready
    }

    $hint = "変換中も検索できます。途中でやめるときは［中止］を押してください（次回、続きから再開できます）。"
    if (!$ready -and !(isConverting)) {
        $hint = "変換するフォルダを追加して、チェックを付けてください。"
    } elseif ($state -and $state.Failed -gt 0 -and !(isConverting)) {
        $hint = "前回失敗したファイルがあります。変換を始めるときに、再変換するかを選べます。" + $hint
    }
    $ui.ConvertHint.Text = $hint
}

function refreshConversionState {
    try {
        $state = getConversionState
    } catch {
        # 変換側が書き込んでいる瞬間などは次の機会に読み直す
        return
    }
    $script:conversionState = $state

    # 失敗したファイルは下の一覧に原因とともに表示する
    $ui.ConversionStateText.Text = if ($state.Pending -gt 0 -and !(isConverting)) { "⏸ 前回の変換が中断しています（残り $($state.Pending) 件）" } else { "" }
    $ui.IndexTabHeader.Text = if ($state.Failed -gt 0) { "⚠ 1 インデックス作成" } else { "1 インデックス作成" }
    updateFailedList $state
    updateIndexSummaryText
    updateConvertButton
}

function updateFailedList {
    # 変換に失敗したファイルと原因（変換一覧のエラー列）を一覧に表示する
    param (
        $state  # getConversionState の結果
    )

    $folderPaths = @{}  # インデックス名 → 変換対象フォルダ（大文字・小文字を区別しない）
    foreach ($folder in $state.Folders) {
        if ($folder.Name) {
            $folderPaths[$folder.Name] = $folder.Path
        }
    }

    $rows = New-Object System.Collections.ArrayList
    foreach ($status in $state.FailedRows) {
        $row = New-Object WinGrep.FailRow
        $row.RelPath = $status.相対パス
        $row.Reason = if ($status.エラー) { $status.エラー } else { "（原因は記録されていません）" }
        $converted = [datetime]::MinValue
        if ([datetime]::TryParseExact([string]$status.変換日時, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$converted)) {
            $row.ConvertedText = formatTime $converted
        }
        $parts = splitIndexRelPath $status.相対パス
        if ($folderPaths.ContainsKey($parts.Name)) {
            $row.SourcePath = Join-Path $folderPaths[$parts.Name] $parts.Rest
        }
        [void]$rows.Add($row)
    }

    $ui.FailedGrid.ItemsSource = $rows
    $ui.FailedHeading.Text = "⚠ 変換に失敗したファイル $($rows.Count) 件"
    $ui.FailedPanel.Visibility = if ($rows.Count -gt 0) { "Visible" } else { "Collapsed" }
}

function openFailedFileFolder {
    # 失敗したファイルの場所をエクスプローラーで開く（ファイルを選択した状態）
    $row = $ui.FailedGrid.SelectedItem
    if ($null -eq $row) {
        return
    }
    if (!$row.SourcePath) {
        setStatus "元のファイルの場所が分かりません（変換一覧に変換対象フォルダの記録がありません）：$($row.RelPath)"
        return
    }
    if (Test-Path -LiteralPath $row.SourcePath -PathType Leaf) {
        Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"$($row.SourcePath)`""
        return
    }
    $dir = Split-Path $row.SourcePath -Parent
    if (Test-Path -LiteralPath $dir -PathType Container) {
        Start-Process -FilePath "explorer.exe" -ArgumentList "`"${dir}`""
        setStatus "ファイルが見つからないため、フォルダを開きました（移動・削除された可能性があります）：$($row.SourcePath)"
        return
    }
    setStatus "ファイルが見つかりません（移動・削除された可能性があります）：$($row.SourcePath)"
}

function updateIndexSummaryText {
    $summary = $script:indexSummary
    if ($null -eq $summary) {
        $ui.IndexSummaryText.Text = "確認中…"
        return
    }
    if ($summary["Count"] -eq 0) {
        $ui.IndexSummaryText.Text = "まだインデックスがありません。"
        return
    }
    $text = "TSV $($summary['Count'].ToString('N0')) 件 ・ 最終変換 $(formatTime $summary['LastWrite'])"
    $state = $script:conversionState
    if ($state -and $state.Done -gt 0) {
        $text = "変換済み $($state.Done.ToString('N0')) ファイル（$text）"
    }
    $ui.IndexSummaryText.Text = $text
}

function refreshIndexSummary {
    # TSV の件数は数えるのに時間がかかることがあるため、別スレッドで数える
    if ($script:summaryRunning) {
        $script:summaryAgain = $true
        return
    }
    $script:summaryRunning = $true
    $script:summaryAgain = $false
    $folders = @(getIndexFolders)
    startJob {
        param ($commonPath, $folders)
        . $commonPath
        getIndexSummary $folders
    } @(${commonPath}, $folders) {
        param ($output, $errorText)
        $script:summaryRunning = $false
        if ($output -and $output.Count -gt 0) {
            $script:indexSummary = $output[0]
        }
        updateIndexSummaryText
        updateSearchTarget
        if ($script:summaryAgain) {
            refreshIndexSummary
        }
    }
}

# ---- 変換の起動と進み具合 ----

function getConversionProgress {
    # 変換一覧から、since 以降に起動した変換の進み具合を返す
    param (
        [datetime]$since
    )

    $progress = @{ Scanned = $false; Processed = 0; Failed = 0; Remaining = 0; Current = "" }
    if (!(Test-Path -LiteralPath ${statusFile})) {
        return $progress
    }
    $sinceSecond = $since.AddTicks(-($since.Ticks % [timespan]::TicksPerSecond))
    # 変換対象の検索が終わると、変換一覧が書き直される
    $progress.Scanned = (Get-Item -LiteralPath ${statusFile}).LastWriteTime -ge $sinceSecond
    if (!$progress.Scanned) {
        return $progress
    }

    $status = readStatusFile
    foreach ($row in $status.Rows.Values) {
        if ($row.状態 -eq ${stateNew}) {
            $progress.Remaining++
            continue
        }
        $converted = [datetime]::MinValue
        if ($row.変換日時 -and [datetime]::TryParseExact($row.変換日時, "yyyy/MM/dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$converted) -and $converted -ge $sinceSecond) {
            $progress.Processed++
            if ($row.状態 -eq ${stateFailed}) {
                $progress.Failed++
            }
        }
    }

    $converting = (readTextShared ${convertingFile}).Split([char[]]"`r`n", [System.StringSplitOptions]::RemoveEmptyEntries)
    if ($converting.Count -gt 0 -and $converting[0].Contains("`t")) {
        $progress.Current = $converting[0].Split("`t", 2)[1]
    }
    return $progress
}

function findRunningConversion {
    # このツールの変換（office_to_tsv.ps1）が実行中なら、そのプロセスを返す（画面を閉じて開き直した場合など）
    $script = "${PSScriptRoot}\office_to_tsv.ps1"
    foreach ($process in @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" -ErrorAction SilentlyContinue)) {
        if ($process.CommandLine -and $process.CommandLine.IndexOf($script, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            try {
                return Get-Process -Id $process.ProcessId -ErrorAction Stop
            } catch {
            }
        }
    }
    return $null
}

function showConversionPanel {
    $ui.ConvertProgressPanel.Visibility = "Visible"
    $ui.ConvertProgress.IsIndeterminate = $true
    $ui.ConvertProgressText.Text = "変換の準備をしています…"
    $ui.ConvertProgressEta.Text = ""
    $ui.ConvertProgressDetail.Text = "変換対象のファイルを確認しています。"
    $ui.ConvertStopButton.Visibility = "Visible"
    $ui.ConvertStopButton.IsEnabled = $true
    $ui.ConvertLogButton.Visibility = "Collapsed"
    $taskbar.ProgressState = "Indeterminate"
}

function startConversion {
    if (isConverting) {
        return
    }
    $existing = findRunningConversion
    if ($existing) {
        adoptConversion $existing
        setStatus "実行中の変換があるため、その進み具合を表示します"
        return
    }

    # 前回失敗し、その後更新されていないファイルを再変換するか聞く
    $retryFailed = $false
    $state = $script:conversionState
    if ($state -and $state.Failed -gt 0) {
        $answer = showMessage ("前回変換に失敗し、その後更新されていないファイルが $($state.Failed) 件あります（パスワード付きなど）。`n`n" +
            "これらも再変換しますか？`n（「いいえ」の場合はスキップして、新しいファイル・更新されたファイルだけを変換します）") "YesNoCancel" "Question" "No"
        if ($answer -eq "Cancel") {
            return
        }
        $retryFailed = $answer -eq "Yes"
    }

    saveTargets
    $arguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"${PSScriptRoot}\office_to_tsv.ps1`""
    if ($retryFailed) {
        $arguments += " -RetryFailed"
    }
    $script:convertStart = Get-Date
    $script:convertRate = $null
    $script:convertProcess = Start-Process -FilePath "powershell.exe" -ArgumentList $arguments -WorkingDirectory ${rootDir} -WindowStyle Hidden -PassThru
    # PowerShell 5.1 では、起動直後にハンドルを取っておかないと終了コードを取得できないことがある
    $null = $script:convertProcess.Handle
    $script:convertAdopted = $false

    showConversionPanel
    setStatus "変換を開始しました"
    updateConvertButton
    updateKillBadge
    $script:convertTimer.Start()
}

function adoptConversion {
    # 画面の外で起動された（または前回の画面で起動した）変換の進み具合を表示する
    param (
        [System.Diagnostics.Process]$process
    )

    $script:convertProcess = $process
    try {
        $null = $process.Handle
    } catch {
    }
    $script:convertStart = $process.StartTime
    $script:convertRate = $null
    $script:convertAdopted = $true
    showConversionPanel
    updateConvertButton
    $script:convertTimer.Start()
}

function stopConversion {
    if (!(isConverting)) {
        return
    }
    $answer = showMessage "変換を中止しますか？`n変換中のファイルが終わったところで止まります。次回は続きから再開できます。" "YesNo" "Question" "No"
    if ($answer -ne "Yes") {
        return
    }
    [System.IO.File]::WriteAllText(${stopRequestFile}, "", ${utf8Bom})
    $ui.ConvertStopButton.IsEnabled = $false
    $ui.ConvertProgressDetail.Text = "中止しています…（変換中のファイルが終わるまでお待ちください）"
    setStatus "変換の中止を要求しました"
}

function updateConversionProgress {
    if (!(isConverting)) {
        finishConversion
        return
    }

    try {
        $progress = getConversionProgress $script:convertStart
    } catch {
        return
    }
    $stopping = !$ui.ConvertStopButton.IsEnabled

    $total = $progress.Processed + $progress.Remaining
    if (!$progress.Scanned) {
        $ui.ConvertProgress.IsIndeterminate = $true
        $ui.ConvertProgressText.Text = "変換対象のファイルを確認しています…"
        $taskbar.ProgressState = "Indeterminate"
        return
    }
    if ($progress.Processed -eq 0) {
        $ui.ConvertProgress.IsIndeterminate = $true
        $ui.ConvertProgressText.Text = if ($progress.Remaining -gt 0) { "$($progress.Remaining) 件のファイルを変換します" } else { "変換が必要なファイルを確認しています…" }
        if (!$stopping) {
            $ui.ConvertProgressDetail.Text = if ($progress.Current) { "変換中のファイル：$($progress.Current)" } else { "" }
        }
        $taskbar.ProgressState = "Indeterminate"
        return
    }

    $ratio = if ($total -gt 0) { $progress.Processed / $total } else { 1 }
    $ui.ConvertProgress.IsIndeterminate = $false
    $ui.ConvertProgress.Value = $ratio
    $taskbar.ProgressState = if ($progress.Failed -gt 0) { "Paused" } else { "Normal" }
    $taskbar.ProgressValue = $ratio

    $text = "変換中… $($progress.Processed.ToString('N0')) / $($total.ToString('N0')) 件"
    if ($progress.Failed -gt 0) {
        $text += "（失敗 $($progress.Failed) 件）"
    }
    $ui.ConvertProgressText.Text = $text

    # 残り時間の目安（最初の1件が終わってからの速さで計算する）
    $now = Get-Date
    if ($null -eq $script:convertRate) {
        $script:convertRate = @{ Time = $now; Processed = $progress.Processed }
    }
    $done = $progress.Processed - $script:convertRate.Processed
    if ($progress.Remaining -eq 0) {
        $ui.ConvertProgressEta.Text = ""
    } elseif ($done -gt 0) {
        $seconds = ($now - $script:convertRate.Time).TotalSeconds / $done * $progress.Remaining
        $ui.ConvertProgressEta.Text = if ($seconds -lt 60) { "残り 1 分未満" } else { "残り約 $([math]::Ceiling($seconds / 60)) 分" }
    }
    if (!$stopping) {
        $ui.ConvertProgressDetail.Text = if ($progress.Current) { "変換中のファイル：$($progress.Current)" } else { "" }
    }
}

function finishConversion {
    $script:convertTimer.Stop()
    $taskbar.ProgressState = "None"
    $exitCode = $null
    try {
        $script:convertProcess.WaitForExit()
        $exitCode = $script:convertProcess.ExitCode
    } catch {
    }
    $progress = $null
    try {
        $progress = getConversionProgress $script:convertStart
    } catch {
    }

    $counts = ""
    if ($progress -and $progress.Processed -gt 0) {
        $counts = "成功 $($progress.Processed - $progress.Failed) 件 / 失敗 $($progress.Failed) 件"
        if ($progress.Remaining -gt 0) {
            $counts += " / 残り $($progress.Remaining) 件"
        }
        $ui.ConvertProgress.IsIndeterminate = $false
        $ui.ConvertProgress.Value = $progress.Processed / ($progress.Processed + $progress.Remaining)
    } else {
        $ui.ConvertProgress.IsIndeterminate = $false
        $ui.ConvertProgress.Value = 0
    }

    if ($exitCode -eq 1) {
        # 変換を続けられないエラー（変換対象フォルダが無い など）
        $message = (readTextShared ${convertErrorFile}).Trim()
        if ($message -eq "") {
            $message = "詳しくはログを確認してください。"
        }
        $ui.ConvertProgressText.Text = "変換できませんでした"
        $ui.ConvertProgressDetail.Text = $message
        setStatus "変換できませんでした：$message"
        showMessage "変換できませんでした。`n`n$message" "OK" "Error" | Out-Null
    } elseif ($exitCode -eq 2) {
        $ui.ConvertProgressText.Text = if ($counts) { "変換を中止しました（$counts）" } else { "変換を中止しました" }
        $ui.ConvertProgressDetail.Text = "次回は続きから再開できます。"
        setStatus $ui.ConvertProgressText.Text
    } elseif ($progress -and $progress.Processed -gt 0) {
        $ui.ConvertProgressText.Text = "変換が終わりました（$counts）"
        $ui.ConvertProgressDetail.Text = if ($progress.Failed -gt 0) { "失敗したファイルと原因は「変換に失敗したファイル」の一覧で確認できます。" } else { "" }
        setStatus $ui.ConvertProgressText.Text
    } else {
        $ui.ConvertProgressText.Text = "変換が必要なファイルはありませんでした"
        $ui.ConvertProgressDetail.Text = ""
        setStatus $ui.ConvertProgressText.Text
    }
    $ui.ConvertProgressEta.Text = ""
    $ui.ConvertStopButton.Visibility = "Collapsed"
    $ui.ConvertLogButton.Visibility = if (Test-Path -LiteralPath ${convertLogFile}) { "Visible" } else { "Collapsed" }

    $script:convertProcess = $null
    $script:sourceFolderMaps = @{}
    refreshConversionState
    refreshIndexSummary
    loadIndexTree  # 新しいインデックス・フォルダをツリーに出す
    updateKillBadge
}
$script:convertTimer = newTimer 1000 { safe { updateConversionProgress } }

# ---- イベント ----

$ui.NewFolderBox.Add_TextChanged({
    $ui.NewFolderPlaceholder.Visibility = if ($ui.NewFolderBox.Text -eq "") { "Visible" } else { "Collapsed" }
})
$ui.NewFolderBox.Add_KeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Return") {
        safe {
            addTargetFolder $ui.NewFolderBox.Text
            $ui.NewFolderBox.Text = ""
        }
        $e.Handled = $true
    }
})
$ui.AddFolderButton.Add_Click({
    safe {
        if ($ui.NewFolderBox.Text.Trim() -eq "") {
            $path = selectFolder "変換する Office ファイルのあるフォルダを選んでください" ""
            if ($path) {
                addTargetFolder $path
            }
        } else {
            addTargetFolder $ui.NewFolderBox.Text
            $ui.NewFolderBox.Text = ""
        }
    }
})
$ui.BrowseFolderButton.Add_Click({
    safe {
        $initial = if ($ui.TargetList.SelectedItem) { $ui.TargetList.SelectedItem.Path } else { "" }
        $path = selectFolder "変換する Office ファイルのあるフォルダを選んでください" $initial
        if ($path) {
            addTargetFolder $path
        }
    }
})
foreach ($control in @($ui.NewFolderBox, $ui.TargetList)) {
    $control.Add_PreviewDragOver({ onFolderDragOver @args })
    $control.Add_PreviewDrop({
        param ($sender, $e)
        safe {
            foreach ($folder in (getDroppedFolders $e)) {
                addTargetFolder $folder
            }
        }
        $e.Handled = $true
    })
}
$ui.TargetList.Add_SelectionChanged({ updateTargetView })
$ui.TargetList.Add_KeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Delete") {
        safe { removeTargetFolder }
    }
})
$ui.RemoveFolderButton.Add_Click({ safe { removeTargetFolder } })
$ui.ChangeFolderPathButton.Add_Click({ safe { changeTargetFolderPath } })
$ui.FailedGrid.Add_MouseDoubleClick({
    param ($sender, $e)
    # 行の上でのダブルクリックだけを対象にする（列見出し・スクロールバーは除く）
    $element = $e.OriginalSource
    while ($element -and !($element -is [System.Windows.Controls.DataGridRow])) {
        if ($element -is [System.Windows.Controls.Primitives.DataGridColumnHeader] -or $element -is [System.Windows.Controls.Primitives.ScrollBar]) {
            return
        }
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }
    if ($element) {
        safe { openFailedFileFolder }
    }
})
$ui.ConvertButton.Add_Click({ safe { startConversion } })
$ui.ConvertStopButton.Add_Click({ safe { stopConversion } })
$ui.ConvertLogButton.Add_Click({
    safe {
        if (Test-Path -LiteralPath ${convertLogFile}) {
            Invoke-Item -LiteralPath ${convertLogFile}
        }
    }
})

# ============================================================================
# ［2 検索］
# ============================================================================

$script:hitRows = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.ResultGrid.ItemsSource = $script:hitRows
$script:hitView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:hitRows)
$script:search = $null
$script:lastSearch = $null
$script:sourceFolderMaps = @{}  # インデックスのフォルダ → インデックス名と変換対象フォルダの対応（getSourceLocation のキャッシュ）
$script:filterText = ""

# 別スレッドで実行する検索（結果は $shared.Queue に少しずつ入れる）
${searchScript} = {
    param ($commonPath, $word, $simpleMatch, $folders, $limit, $shared)
    try {
        . $commonPath
        $index = getIndexTsvFiles $folders
        $shared.Folders = $index.Folders
        $shared.Total = $index.Files.Count
        $shared.IndexTotal = $index.Files.Count
        # 検索条件（大文字・小文字の区別・対象ファイル）は startSearch が $shared に入れる
        $result = searchIndex $word $index.Files $simpleMatch $limit 50 -caseSensitive $shared.CaseSensitive -fileFilter $shared.FileFilter {
            param ($done, $total, $newHits)
            foreach ($hit in $newHits) {
                $shared.Queue.Enqueue($hit)
            }
            $shared.Total = $total
            $shared.Done = $done
        } { $shared.Stop }
        $shared.Total = $result.Total
        $shared.Truncated = $result.Truncated
        $shared.Cancelled = $result.Cancelled
    } catch {
        $shared.Error = $_.Exception.Message
    } finally {
        $shared.Finished = $true
    }
}

function getWordText {
    return $ui.WordBox.Text.Trim()
}

function getSearchOptionFromUi {
    # 画面の検索条件を readSearchOption と同じ形で返す
    return @{
        UseRegex      = [bool]$ui.RegexCheck.IsChecked
        CaseSensitive = [bool]$ui.CaseCheck.IsChecked
        FileFilter    = $ui.FileFilterBox.Text.Trim()
    }
}

function setSearchOptionToUi {
    param (
        [hashtable]$option
    )

    $ui.RegexCheck.IsChecked = [bool]$option.UseRegex
    $ui.CaseCheck.IsChecked = [bool]$option.CaseSensitive
    $ui.FileFilterBox.Text = [string]$option.FileFilter
}

function describeSearchOption {
    # 既定から変えた検索条件を「大文字と小文字を区別・対象ファイル：*.xlsx」のように返す（無ければ空）
    param (
        [hashtable]$option
    )

    $items = @()
    if ($option.CaseSensitive) {
        $items += "大文字と小文字を区別"
    }
    if ($option.FileFilter) {
        $items += "対象ファイル：$($option.FileFilter)"
    }
    return ($items -join "・")
}

function updateWordNotice {
    $word = getWordText
    if ($ui.RegexCheck.IsChecked -and $word -ne "" -and !(isValidRegex $word)) {
        $ui.WordNotice.Text = "正規表現として不正なため、文字どおり検索します。"
        $ui.WordNotice.Visibility = "Visible"
    } else {
        $ui.WordNotice.Visibility = "Collapsed"
    }
    updateSearchButton
}

function updateSearchButton {
    if ($script:search) {
        $ui.SearchButton.Content = "中止"
        $ui.SearchButton.IsEnabled = !$script:search.Shared.Stop
        return
    }
    $ui.SearchButton.Content = "検索"
    $noIndex = $script:indexSummary -and $script:indexSummary["Count"] -eq 0
    $ui.SearchButton.IsEnabled = (getWordText) -ne "" -and !$noIndex -and @(getSearchTargets).Count -gt 0
}

function toDisplayFolder {
    param (
        [string]$folder
    )

    try {
        $full = (Resolve-Path -LiteralPath $folder -ErrorAction Stop).ProviderPath.TrimEnd("\")
        if ($full -eq ${indexDir}.TrimEnd("\")) {
            return "work\index"
        }
    } catch {
    }
    return $folder
}

function updateSearchTarget {
    $targets = @(getSearchTargets)
    $summary = $script:indexSummary
    if ($summary -and $summary["Count"] -eq 0) {
        $names = @(getIndexFolders | ForEach-Object { toDisplayFolder $_ }) -join "、"
        $ui.SearchTargetText.Text = "検索対象：${names}（TSV がありません。先にインデックスを作成してください）"
    } elseif ($targets.Count -eq 0) {
        $ui.SearchTargetText.Text = "検索対象：なし（左の一覧で、検索するフォルダにチェックを付けてください）"
    } elseif (!(isAllIndexChecked)) {
        $ui.SearchTargetText.Text = "検索対象：$(describeSearchTargets $targets)"
    } elseif ($null -eq $summary) {
        $ui.SearchTargetText.Text = "検索対象：すべて（確認中…）"
    } else {
        $ui.SearchTargetText.Text = "検索対象：すべて（TSV $($summary['Count'].ToString('N0')) 件 ・ 最終変換 $(formatTime $summary['LastWrite'])）"
    }
    $ui.SearchTargetText.ToolTip = $ui.SearchTargetText.Text
    $ui.GoIndexTabButton.Visibility = if ($summary -and $summary["Count"] -eq 0) { "Visible" } else { "Collapsed" }
    updateSearchButton
}

function loadHistory {
    $text = $ui.WordBox.Text
    $ui.WordBox.ItemsSource = [string[]]@(readSearchHistory)
    $ui.WordBox.Text = $text
}

function startSearch {
    if ($script:search) {
        cancelSearch
        return
    }
    $word = getWordText
    if ($word -eq "") {
        setStatus "検索ワードを入力してください。"
        return
    }

    $option = getSearchOptionFromUi
    writeSearchOption $option
    $useRegex = $option.UseRegex
    # 一致箇所の強調にも、検索と同じ正規表現を使う
    $searchRegex = newSearchRegex $word (!$useRegex) $option.CaseSensitive
    $simpleMatch = $searchRegex.SimpleMatch
    $pattern = $searchRegex.Regex

    $ui.FilterBox.Text = ""
    $script:filterText = ""
    $script:hitView.Filter = $null
    $script:hitRows.Clear()
    $ui.DetailPanel.Visibility = "Collapsed"
    # 検索対象ツリーでチェックしたフォルダだけを検索する（結果の相対パスは、インデックスのフォルダからのまま）
    $folders = @(getSearchTargets)
    if ($folders.Count -eq 0) {
        setStatus "検索するフォルダに、左の「検索対象」でチェックを付けてください。"
        return
    }
    $ui.IndexColumn.Visibility = if (@($folders | ForEach-Object { $_.Root } | Sort-Object -Unique).Count -gt 1) { "Visible" } else { "Collapsed" }

    $shared = [hashtable]::Synchronized(@{
        Queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
        Stop = $false; Finished = $false; Done = 0; Total = -1; IndexTotal = -1; Folders = $null
        Truncated = $false; Cancelled = $false; Error = $null
        CaseSensitive = $option.CaseSensitive; FileFilter = $option.FileFilter
    })
    $ps = [powershell]::Create()
    [void]$ps.AddScript(${searchScript}.ToString())
    foreach ($argument in @(${commonPath}, $word, $simpleMatch, $folders, ${searchLimit}, $shared)) {
        [void]$ps.AddArgument($argument)
    }
    $script:search = @{
        PS = $ps; Handle = $ps.BeginInvoke(); Shared = $shared
        Word = $word; Pattern = $pattern; SimpleMatch = $simpleMatch; UseRegex = $useRegex; Option = $option; Start = Get-Date
    }

    $ui.SearchProgress.Visibility = "Visible"
    $ui.SearchProgress.IsIndeterminate = $true
    $ui.SummaryText.Text = "検索中…"
    $taskbar.ProgressState = "Indeterminate"
    setStatus "検索しています：${word}"
    updateSearchButton
    $script:searchTimer.Start()
}

function cancelSearch {
    if ($script:search) {
        $script:search.Shared.Stop = $true
        $ui.SummaryText.Text = "中止しています…"
        updateSearchButton
    }
}

function pumpSearch {
    # 検索スレッドの結果を表に移し、進み具合を表示する
    $s = $script:search
    if ($null -eq $s) {
        $script:searchTimer.Stop()
        return
    }
    $shared = $s.Shared
    $hit = $null
    $added = 0
    while ($added -lt 3000 -and $shared.Queue.TryDequeue([ref]$hit)) {
        $script:hitRows.Add([WinGrep.HitRow]::Create((Split-Path $hit.Root -Leaf), $hit.Root, $hit.RelPath, $hit.RelDir, $hit.FileName,
                $hit.Book, $hit.Location, [int]$hit.LineNumber, $hit.Line, $s.Word, $s.Pattern))
        $added++
    }

    if ($shared.Total -gt 0) {
        $ratio = $shared.Done / $shared.Total
        $ui.SearchProgress.IsIndeterminate = $false
        $ui.SearchProgress.Value = $ratio
        $taskbar.ProgressState = "Normal"
        $taskbar.ProgressValue = $ratio
        if (!$shared.Stop) {
            $ui.SummaryText.Text = "検索中… $($shared.Done.ToString('N0')) / $($shared.Total.ToString('N0')) ファイル（$($script:hitRows.Count.ToString('N0')) 件）"
        }
    } elseif ($shared.Total -lt 0 -and !$shared.Stop) {
        $ui.SummaryText.Text = "検索対象のファイルを確認しています…"
    }

    if ($shared.Finished -and $shared.Queue.IsEmpty) {
        finishSearch
    }
}

function finishSearch {
    $s = $script:search
    $script:search = $null
    $script:searchTimer.Stop()
    try {
        [void]$s.PS.EndInvoke($s.Handle)
    } catch {
    }
    $s.PS.Dispose()

    $shared = $s.Shared
    $ui.SearchProgress.Visibility = "Collapsed"
    $taskbar.ProgressState = if (isConverting) { $taskbar.ProgressState } else { "None" }
    $script:lastSearch = $s
    $seconds = ((Get-Date) - $s.Start).TotalSeconds
    updateSearchButton

    if ($shared.Error) {
        $ui.SummaryText.Text = "検索できませんでした"
        setStatus "検索できませんでした：$($shared.Error)"
        return
    }

    $count = $script:hitRows.Count
    $files = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $script:hitRows) {
        [void]$files.Add("$($row.Root)\$($row.RelDir)\$($row.Book)")
    }

    if ($shared.IndexTotal -gt 0 -and $shared.Total -eq 0) {
        $ui.SummaryText.Text = "対象ファイル（$($s.Option.FileFilter)）に一致するファイルがありません。"
    } elseif ($shared.Total -eq 0) {
        $ui.SummaryText.Text = "検索対象の TSV がありません。先にインデックスを作成してください。"
    } elseif ($count -eq 0) {
        $text = "見つかりませんでした。"
        if (!$s.UseRegex -and $s.Word -match '[\\()\[\]{}.*+?^$|]') {
            $text += "（正規表現として探す場合は［正規表現を使う］をオンにしてください）"
        } elseif (describeSearchOption $s.Option) {
            $text += "（検索条件：$(describeSearchOption $s.Option)）"
        }
        $ui.SummaryText.Text = $text
    } else {
        $ui.SummaryText.Text = "$($count.ToString('N0')) 件（$($files.Count.ToString('N0')) ファイル） ・ $($seconds.ToString('0.0')) 秒"
    }

    $status = "検索しました（$($s.Word)：$($count.ToString('N0')) 件）"
    if (describeSearchOption $s.Option) {
        $status = "検索しました（$($s.Word)：$($count.ToString('N0')) 件　条件：$(describeSearchOption $s.Option)）"
    }
    if ($shared.Truncated) {
        $status = "$(${searchLimit}.ToString('N0')) 件を超えたため打ち切りました。ワードを絞り込んでください。"
    } elseif ($shared.Cancelled) {
        $status = "中止しました（$($count.ToString('N0')) 件まで表示）"
    }
    $missing = @($shared.Folders | Where-Object { !$_.Exists } | ForEach-Object { $_.Path })
    if ($missing.Count -gt 0) {
        $status += "　見つからない検索対象フォルダ：$($missing -join '、')"
    }
    if (isConverting) {
        $status += "　変換中のため、作成途中のインデックスを検索しています。"
    }
    setStatus $status

    [void](addSearchHistory $s.Word)
    loadHistory
}

$script:searchTimer = newTimer 100 { safe { pumpSearch } }

# ---- 絞り込み・選択行の詳細 ----

function applyFilter {
    $script:filterText = $ui.FilterBox.Text.Trim()
    if ($script:filterText -eq "") {
        $script:hitView.Filter = $null
    } else {
        $script:hitView.Filter = [Predicate[object]] { param ($row) $row.Contains($script:filterText) }
    }
    if ($script:lastSearch -and !$script:search -and $script:hitRows.Count -gt 0) {
        $shown = 0
        foreach ($row in $script:hitView) {
            $shown++
        }
        if ($script:filterText -eq "") {
            finishSummaryText
        } else {
            $ui.SummaryText.Text = "$($script:hitRows.Count.ToString('N0')) 件中 $($shown.ToString('N0')) 件を表示"
        }
    }
}

function finishSummaryText {
    $files = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $script:hitRows) {
        [void]$files.Add("$($row.Root)\$($row.RelDir)\$($row.Book)")
    }
    $ui.SummaryText.Text = "$($script:hitRows.Count.ToString('N0')) 件（$($files.Count.ToString('N0')) ファイル）"
}

$script:filterTimer = newTimer 300 {
    $script:filterTimer.Stop()
    safe { applyFilter }
}

function getViewRows {
    # 表示中（絞り込み・並べ替え後）の行
    $rows = New-Object System.Collections.ArrayList
    foreach ($row in $script:hitView) {
        [void]$rows.Add($row)
    }
    return , $rows.ToArray()
}

function getSelectedRows {
    # 選択行を表示の順に並べて返す
    $rows = New-Object System.Collections.ArrayList
    foreach ($row in $ui.ResultGrid.SelectedItems) {
        [void]$rows.Add($row)
    }
    return , @($rows.ToArray() | Sort-Object { $ui.ResultGrid.Items.IndexOf($_) })
}

function showDetail {
    $row = $ui.ResultGrid.SelectedItem
    if ($null -eq $row) {
        $ui.DetailPanel.Visibility = "Collapsed"
        return
    }
    $path = if ($row.RelDir) { "$($row.RelDir)\$($row.Book)" } else { $row.Book }
    $place = if ($row.MatchCell) { "セル $($row.MatchCell)" } else { "$($row.LineNumber) 行目" }
    $title = "${path} ・ $($row.Location) ・ ${place}"
    $ui.DetailTitle.Text = $title
    $ui.DetailTitle.ToolTip = $title
    $ui.OpenButton.Content = if ($row.IsExcel) { "Excel で開く" } else { "開く" }

    # 前後の行をインデックスのTSVから読む（読めなければ選択行だけを出す）
    $context = @(readTsvContext ([System.IO.Path]::Combine($row.Root, $row.RelPath)) $row.LineNumber ${previewLines} ${previewLines})
    $table = $row.BuildPreview([int[]]@($context | ForEach-Object { $_.LineNumber }), [string[]]@($context | ForEach-Object { $_.Line }))
    $ui.PreviewHeader.ItemsSource = $table.Columns
    $ui.PreviewRows.ItemsSource = $table.Rows
    $script:previewTable = $table
    $ui.DetailPanel.Visibility = "Visible"

    # 一致したセルが見えるよう横にスクロールする（左端から見えていればそのまま）
    $ui.PreviewScroll.UpdateLayout()
    $offset = 0
    if ($table.HitOffset + $table.HitWidth -gt $ui.PreviewScroll.ViewportWidth) {
        $offset = [math]::Max(0, $table.HitOffset - 120)
    }
    $ui.PreviewScroll.ScrollToHorizontalOffset($offset)
}

function getPreviewCell {
    # マウスの下（またはイベントの発生元）のプレビューのセル。セルの上でなければ $null
    param (
        $source
    )

    $element = $source
    while ($element) {
        if ($element -is [System.Windows.FrameworkElement] -and $element.DataContext -is [WinGrep.PreviewCell]) {
            return $element.DataContext
        }
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }
    return $null
}

function copyPreviewSelection {
    # プレビューで選んだセルの値をクリップボードに入れる（1 セルならその値のまま、複数ならタブ区切り）
    if ($null -eq $script:previewTable -or !$script:previewTable.HasSelection) {
        setStatus "プレビューでコピーするセルをクリックしてください（Shift＋クリック・ドラッグで複数選べます）。"
        return
    }
    $text = $script:previewTable.GetSelectionText()
    if ($text -eq "") {
        [System.Windows.Clipboard]::Clear()
    } else {
        [System.Windows.Clipboard]::SetText($text)
    }
    $count = $script:previewTable.SelectedCount
    if ($count -le 1) {
        setStatus "セルの値をコピーしました：$(toStatusText $text)"
    } else {
        setStatus "${count} 個のセルをコピーしました（Excel に貼り付けると、元の位置に並びます）"
    }
}

function toStatusText {
    # ステータスに出す短い文字列（改行・タブはスペースにし、長ければ末尾を省略）
    param (
        [string]$text
    )

    $text = ($text -replace "[`r`n`t]+", " ")
    if ($text.Length -gt 40) {
        return $text.Substring(0, 40) + "…"
    }
    return $text
}

# ---- 元のファイルを開く・コピー・出力 ----

function getSourcePath {
    # 元のファイルのパス（ファイルがあるかは確かめない）。元の場所が分からなければ $null
    param (
        $row
    )

    return resolveSourcePath $row $script:sourceFolderMaps
}

function getExistingFolder {
    # path の上のフォルダのうち、存在する最も深いフォルダ（フォルダ選択の初期位置）。無ければ空
    param (
        [string]$path
    )

    $dir = Split-Path $path -Parent
    while ($dir) {
        if (Test-Path -LiteralPath $dir -PathType Container) {
            return $dir
        }
        $dir = Split-Path $dir -Parent
    }
    return ""
}

function findSourceFile {
    # 元のファイルのパスを返す。見つからなければ、元のファイルのあるフォルダを選んでもらって探し、
    # 見つかればそのインデックスの元のフォルダ（今の置き場所）として設定に記録する
    # （インデックス名に対して 1 か所を記録するため、同じインデックスのほかのファイルも次からそのまま開ける）。
    # 見つからない・選ばなかった場合は $null
    param (
        $row
    )

    $location = getSourceLocation $row $script:sourceFolderMaps
    $relPath = if ($location.Rest) { "$($location.Rest)\$($row.Book)" } else { $row.Book }
    if ($location.Known) {
        $path = joinSourcePath $location.Folder $location.Rest $row.Book
        if (Test-Path -LiteralPath (toLongPath $path) -PathType Leaf) {
            return $path
        }
        # 記録した場所に無くても、書き方が違うだけで同じ場所を指すパスで開けることがある
        # （ネットワークドライブと UNC パス）。別の PC でドライブの割り当てが違う場合に、聞かずに開けるようにする
        foreach ($alias in @(getFolderPathAliases $location.Folder | Select-Object -Skip 1)) {
            $candidate = joinSourcePath $alias $location.Rest $row.Book
            if (!(Test-Path -LiteralPath (toLongPath $candidate) -PathType Leaf)) {
                continue
            }
            if ($location.Name) {
                setIndexSourceFolder $location.Name $alias
                $script:sourceFolderMaps = @{}
                setStatus "インデックス [$($location.Name)] の元のフォルダを ${alias} に変えました"
            }
            return $candidate
        }
        $message = "元のファイルが見つかりません。`n${path}`n`n" +
                   "インデックスを別の PC に持ってきた場合や、フォルダを移した場合は、今の場所のフォルダを選ぶと開けます。"
        $description = "「$($location.Folder)」に当たるフォルダ（または $($row.Book) のあるフォルダ）を選んでください"
        $initial = getExistingFolder $path
    } else {
        $path = "$($row.Root)\$($row.RelDir)\$($row.Book)"
        $message = "元のファイルの場所が分かりません（インデックス [$($location.Name)] の元のフォルダが記録されていません）。`n${relPath}`n`n" +
                   "元のファイルのあるフォルダを選ぶと開けます。"
        $description = "$($row.Book) のあるフォルダ（またはインデックス [$($location.Name)] の元のフォルダ）を選んでください"
        $initial = ""
    }
    $message += "`n（選んだフォルダはインデックス [$($location.Name)] の元のフォルダとして記録し、同じインデックスのほかのファイルも開けるようにします）`n`nフォルダを選びますか？"

    while ($true) {
        if ((showMessage $message "YesNo" "Question" "Yes") -ne "Yes") {
            setStatus "元のファイルが見つかりません：${path}"
            return $null
        }
        $picked = selectFolder $description $initial
        if (!$picked) {
            setStatus "元のファイルが見つかりません：${path}"
            return $null
        }

        $found = findMovedSource $picked $location.Rest $row.Book
        if ($found) {
            if ($found.Root -and $location.Name) {
                setIndexSourceFolder $location.Name $found.Root
                $script:sourceFolderMaps = @{}
                setStatus "インデックス [$($location.Name)] の元のフォルダを $($found.Root) に変えました"
            } else {
                setStatus "開きました：$($found.Path)"
            }
            return $found.Path
        }
        $message = "選んだフォルダの中に、元のファイルが見つかりませんでした。`n選んだフォルダ：${picked}`n探したファイル：${relPath}`n`n別のフォルダを選びますか？"
        $initial = $picked
    }
}

function openWithShell {
    # ファイルを既定のアプリで開く。開き方（mode）は、エクスプローラーの右クリックメニューと同じ動詞で行う。
    #   読み取り専用 → OpenAsReadOnly、新規 → New（元のファイルを基にした無題の文書。元のファイルを占有しない）
    # その動詞が登録されていない種類のファイルは、そのまま開いて $false を返す
    param (
        [string]$path,
        [string]$mode
    )

    $verb = switch ($mode) {
        ${openModeReadOnly} { "OpenAsReadOnly" }
        ${openModeNew}      { "New" }
        default             { $null }
    }
    if ($verb) {
        # Start-Process は [ ] をワイルドカードとして扱うため使わない。
        # ProcessStartInfo.Verbs は New を一覧に含めないため、実行してみて、無ければ例外で分かる
        $info = New-Object System.Diagnostics.ProcessStartInfo($path)
        $info.UseShellExecute = $true
        $info.Verb = $verb
        try {
            [void][System.Diagnostics.Process]::Start($info)
            return $true
        } catch [System.ComponentModel.Win32Exception] {
            # その動詞が登録されていない
        }
    }
    Invoke-Item -LiteralPath $path
    return (-not $verb)
}

function openInExcel {
    # 表示中の Excel（無ければ新しく起動）でブックを開き、該当シートの該当セルを選択する。
    # 開き方（mode）: 通常 = そのまま開く / 読み取り専用 = ReadOnly で開く /
    #                 新規 = 元のファイルを基にした新しいブック（無題）として開く（元のファイルを占有しない）
    param (
        [string]$path,
        [string]$location,
        [string]$cell,
        [string]$mode = ${openModeNormal}
    )

    $excel = $null
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        # 変換処理がバックグラウンドで使っている Excel は使わない
        if (!$excel.Visible) {
            $excel = $null
        }
    } catch {
        $excel = $null
    }
    if ($null -eq $excel) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        $excel.UserControl = $true
    }

    $book = $null
    if ($mode -eq ${openModeNew}) {
        # 元のファイルをテンプレートとして新しいブックを作る（読み込んだ後は元のファイルを開いたままにしない）
        $book = $excel.Workbooks.Add($path)
    } else {
        # 既に開いているブックがあれば、そのまま使う（同じブックを二重に開けないため。開き方も既に開いたときのまま）
        foreach ($openBook in $excel.Workbooks) {
            if ($openBook.FullName -eq $path) {
                $book = $openBook
                break
            }
        }
        if ($null -eq $book) {
            #   引数: Filename, UpdateLinks, ReadOnly
            $book = $excel.Workbooks.Open($path, [Type]::Missing, ($mode -eq ${openModeReadOnly}))
        }
    }
    $book.Activate()

    # 場所はシート名。名前が同じシートを選ぶ。
    # 以前の版のインデックスは、ファイル名に使えない文字を全角に置き換えてあるため、同じ名前のシートが無ければ
    # 全角に置き換えて一致するシートを選ぶ（`衝突"` と `衝突”` のように、置き換えると重なるシートがあるため、同じ名前を優先する）
    $target = $null
    $sameSafeName = $null
    foreach ($sheet in $book.Worksheets) {
        if ($sheet.Name -eq $location) {
            $target = $sheet
            break
        }
        if ($null -eq $sameSafeName -and (toSafeFileName $sheet.Name) -eq $location) {
            $sameSafeName = $sheet
        }
    }
    if ($null -eq $target) {
        $target = $sameSafeName
    }
    if ($null -ne $target) {
        $target.Activate()
        if ($cell) {
            $target.Range($cell).Select()
        }
    }
    if ($excel.WindowState -eq -4140) {
        # 最小化されていれば元に戻す（xlMinimized → xlNormal）
        $excel.WindowState = -4143
    }
    [void][WinGrep.Native]::SetForegroundWindow([IntPtr][int]$excel.Hwnd)
}

function getOpenMode {
    # ダブルクリック・Enter・［開く］での開き方（［開き方］の選択。${openModes} のいずれか）
    $item = $ui.OpenModeCombo.SelectedItem
    if ($null -ne $item -and (${openModes} -contains $item.Tag)) {
        return [string]$item.Tag
    }
    return ${openModeNormal}
}

function setOpenMode {
    # ［開き方］の選択を設定の値に合わせる（起動時。選んだことにはしないため、設定は保存しない）
    param (
        [string]$mode
    )

    $script:loadingOpenMode = $true
    try {
        foreach ($item in $ui.OpenModeCombo.Items) {
            if ($item.Tag -eq $mode) {
                $ui.OpenModeCombo.SelectedItem = $item
                return
            }
        }
        $ui.OpenModeCombo.SelectedIndex = 0
    } finally {
        $script:loadingOpenMode = $false
    }
}

function updateOpenMenu {
    # 右クリックメニューは3つの開き方をすべて出し、既定の開き方（ダブルクリック・Enter と同じ）に Enter を表示する
    $mode = getOpenMode
    $ui.MenuOpen.InputGestureText         = $(if ($mode -eq ${openModeNormal})   { "Enter" } else { "" })
    $ui.MenuOpenReadOnly.InputGestureText = $(if ($mode -eq ${openModeReadOnly}) { "Enter" } else { "" })
    $ui.MenuOpenNew.InputGestureText      = $(if ($mode -eq ${openModeNew})      { "Enter" } else { "" })
}

function openSource {
    # 選択行の元のファイルを開く。mode で開き方（通常・読み取り専用・新規）を指定する
    param (
        [string]$mode = (getOpenMode)
    )

    $row = $ui.ResultGrid.SelectedItem
    if ($null -eq $row) {
        return
    }
    $path = findSourceFile $row
    if (!$path) {
        return
    }
    $how = switch ($mode) {
        ${openModeReadOnly} { "読み取り専用で開きました" }
        ${openModeNew}      { "新規で開きました" }
        default             { "開きました" }
    }
    $fallback = switch ($mode) {
        ${openModeReadOnly} { "読み取り専用で開けなかったため、元のファイルを開きました" }
        default             { "新規で開けなかったため、元のファイルを開きました" }
    }

    if ($row.IsExcel) {
        setStatus "Excel で開いています：${path}"
        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        try {
            openInExcel $path $row.Location $row.MatchCell $mode
            setStatus "${how}：${path}"
        } catch {
            # Excel を操作できない場合（ダイアログを表示中など）は、ファイルを開くだけにする
            if (openWithShell $path $mode) {
                setStatus "${how}（該当セルへの移動はできませんでした）：${path}"
            } else {
                setStatus "${fallback}（該当セルへの移動はできませんでした）：${path}"
            }
        } finally {
            $window.Cursor = $null
        }
    } elseif (openWithShell $path $mode) {
        setStatus "${how}：${path}"
    } else {
        setStatus "${fallback}：${path}"
    }
}

function openSourceFolder {
    $row = $ui.ResultGrid.SelectedItem
    if ($null -eq $row) {
        return
    }
    $path = findSourceFile $row
    if (!$path) {
        return
    }
    Start-Process -FilePath "explorer.exe" -ArgumentList "/select,`"${path}`""
}

function copySelectedRows {
    $rows = getSelectedRows
    if ($rows.Count -eq 0) {
        return
    }
    $result = toSearchResultLines $rows
    $text = ((@($result.Header) + $result.Lines.ToArray()) -join "`r`n") + "`r`n"
    [System.Windows.Clipboard]::SetText($text)
    setStatus "$($rows.Count) 行をコピーしました（Excel に貼り付けると、元の列の位置に並びます）"
}

function copySourcePath {
    $row = $ui.ResultGrid.SelectedItem
    if ($null -eq $row) {
        return
    }
    $path = getSourcePath $row
    if (!$path) {
        $path = "$($row.Root)\$($row.RelPath)"
    }
    [System.Windows.Clipboard]::SetText($path)
    setStatus "パスをコピーしました：${path}"
}

function exportResults {
    if ($null -eq $script:lastSearch) {
        setStatus "先に検索してください。"
        return
    }
    $rows = getViewRows
    try {
        [System.IO.Directory]::CreateDirectory(${outputDir}) | Out-Null
        $writer = New-Object System.IO.StreamWriter(${resultFile}, $false, ${utf8Bom})
        try {
            writeSearchResult $writer $script:lastSearch.Word $rows
        } finally {
            $writer.Close()
        }
    } catch [System.IO.IOException] {
        setStatus "検索結果.txt に書き込めません。開いているアプリを閉じてから、もう一度出力してください。"
        return
    }
    Invoke-Item -LiteralPath ${resultFile}
    if ($rows.Count -lt $script:hitRows.Count) {
        setStatus "絞り込み後の $($rows.Count.ToString('N0')) 件を検索結果.txt に出力しました"
    } else {
        setStatus "検索結果.txt に出力しました（$($rows.Count.ToString('N0')) 件）"
    }
}

# ---- 検索対象インデックスのツリー ----

$script:indexRoots = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$ui.IndexTree.ItemsSource = $script:indexRoots

function loadIndexTree {
    # 検索対象インデックスのフォルダをツリーに読み込み、保存したチェックなしのフォルダと、読み込み前の展開の状態を戻す
    # （初めて読み込むときは、インデックスのフォルダを展開してインデックスの一覧を見せる）
    $expanded = New-Object 'System.Collections.Generic.List[string]'
    foreach ($node in $script:indexRoots) {
        $node.AddExpanded($expanded)
    }
    $first = $script:indexRoots.Count -eq 0

    $script:indexRoots.Clear()
    foreach ($folder in @(getIndexFolders)) {
        $root = $folder
        $sourcePath = $null
        $map = $null
        if (Test-Path -LiteralPath $folder -PathType Container) {
            $root = (Resolve-Path -LiteralPath $folder).ProviderPath.TrimEnd("\")
            $map = getSourceFolderMap $root
            # インデックス名のフォルダ（…\index\<インデックス名>）を直接指定した場合は、親フォルダの記録から元のフォルダを引く
            $parent = Split-Path $root -Parent
            if ($parent) {
                $parentMap = getSourceFolderMap $parent
                $leaf = Split-Path $root -Leaf
                if ($parentMap.ContainsKey($leaf)) {
                    $sourcePath = $parentMap[$leaf]
                }
            }
        }
        $script:indexRoots.Add([WinGrep.IndexNode]::CreateRoot($root, (toDisplayFolder $folder), $sourcePath, $map))
    }

    foreach ($exclude in @(readSearchExcludes)) {
        foreach ($node in $script:indexRoots) {
            $node.ApplyExclude($exclude.Path, $exclude.Subfolders)
        }
    }
    foreach ($node in $script:indexRoots) {
        if ($first) {
            $node.IsExpanded = $true
            continue
        }
        foreach ($path in $expanded) {
            $found = $node.Find($path)
            if ($found) {
                $found.IsExpanded = $true
            }
        }
    }
    updateSearchTarget
}

function getSearchTargets {
    # 検索対象ツリーでチェックしたフォルダ（WinGrep.SearchTarget の配列。getIndexTsvFiles に渡す）
    $targets = New-Object 'System.Collections.Generic.List[WinGrep.SearchTarget]'
    foreach ($node in $script:indexRoots) {
        $node.AddTargets($targets)
    }
    return $targets.ToArray()
}

function isAllIndexChecked {
    return @($script:indexRoots | Where-Object { $_.IsChecked -ne $true }).Count -eq 0
}

function describeSearchTargets {
    # 検索対象の表示（先頭の 3 件まで）。インデックスのフォルダが複数あるときは、フォルダの表示名を先頭に付ける
    param (
        [object[]]$targets
    )

    $names = @($targets | ForEach-Object {
        $target = $_
        $root = @($script:indexRoots | Where-Object { $_.Root -eq $target.Root })[0]
        $name = if (!$target.RelPath) { $root.Name } elseif ($script:indexRoots.Count -gt 1) { "$($root.Name)\$($target.RelPath)" } else { $target.RelPath }
        if (!$target.Recurse) {
            $name += "（直下のファイル）"
        }
        $name
    })
    if ($names.Count -gt 3) {
        return "$($names[0..2] -join '、') ほか $($names.Count - 3) か所"
    }
    return $names -join "、"
}

function saveSearchExcludes {
    # チェックなしのフォルダを設定に保存する。見つからないインデックスのフォルダ（ネットワークのドライブが切れているなど）の記録は残す
    $excludes = New-Object 'System.Collections.Generic.List[WinGrep.SearchExclude]'
    foreach ($node in $script:indexRoots) {
        $node.AddExcludes($excludes)
    }
    $roots = @($script:indexRoots | Where-Object { $_.Exists } | ForEach-Object { $_.FullPath.TrimEnd("\") })
    $kept = @(readSearchExcludes | Where-Object {
        $path = $_.Path
        @($roots | Where-Object { $path -eq $_ -or $path.StartsWith("$_\", [System.StringComparison]::OrdinalIgnoreCase) }).Count -eq 0
    })
    writeSearchExcludes (@($kept) + @($excludes.ToArray()))
}

function onIndexTreeChecked {
    saveSearchExcludes
    updateSearchTarget
}

function setAllIndexChecked {
    param (
        [bool]$checked
    )

    foreach ($node in $script:indexRoots) {
        $node.SetChecked($checked)
    }
    onIndexTreeChecked
}

# ---- 検索対象インデックスの設定ダイアログ ----

$script:indexDialog = $null

function newIndexListItem {
    param (
        [string]$path
    )

    $item = New-Object System.Windows.Controls.ListBoxItem
    $item.Content = $path
    if (!(Test-Path -LiteralPath $path -PathType Container)) {
        $item.Foreground = ${warnBrush}
        $item.ToolTip = "フォルダが見つかりません（検索時はスキップされます）"
    }
    return $item
}

function updateIndexDialogView {
    $d = $script:indexDialog
    $d.Placeholder.Visibility = if ($d.List.Items.Count -eq 0) { "Visible" } else { "Collapsed" }
}

function addIndexDialogFolder {
    param (
        [string]$path
    )

    $list = $script:indexDialog.List
    $path = normalizeFolderPath $path
    if ($path -eq "") {
        return
    }
    foreach ($existing in $list.Items) {
        # 書き方が違うだけで同じフォルダ（ネットワークドライブと UNC パスなど）も、すでにあるとみなす
        if (testSameFolder ([string]$existing.Content) $path) {
            $list.SelectedItem = $existing
            return
        }
    }
    $list.SelectedIndex = $list.Items.Add((newIndexListItem $path))
    updateIndexDialogView
}

function removeIndexDialogFolder {
    $list = $script:indexDialog.List
    if ($list.SelectedIndex -lt 0) {
        return
    }
    $path = [string]$list.SelectedItem.Content
    $answer = showMessage ("「${path}」を検索対象インデックスの一覧から削除します。`n" +
        "フォルダとその中の TSV は削除しません（検索の対象から外れるだけです）。`n`n削除しますか？") "YesNo" "Question" "No" $script:indexDialog.Window
    if ($answer -ne "Yes") {
        return
    }
    $list.Items.RemoveAt($list.SelectedIndex)
    updateIndexDialogView
}

function moveIndexDialogFolder {
    param (
        [int]$offset
    )

    $list = $script:indexDialog.List
    $from = $list.SelectedIndex
    $to = $from + $offset
    if ($from -lt 0 -or $to -lt 0 -or $to -ge $list.Items.Count) {
        return
    }
    $item = $list.Items[$from]
    $list.Items.RemoveAt($from)
    $list.Items.Insert($to, $item)
    $list.SelectedIndex = $to
}

function showIndexDialog {
    $dialog = loadWindow "$PSScriptRoot\config_gui_index.xaml"
    $dialog.Owner = $window
    $script:indexDialog = @{ Window = $dialog; List = $dialog.FindName("FolderList"); Placeholder = $dialog.FindName("Placeholder") }
    $list = $script:indexDialog.List

    foreach ($path in @(readIndexFolders)) {
        [void]$list.Items.Add((newIndexListItem $path))
    }
    updateIndexDialogView

    $dialog.FindName("AddButton").Add_Click({
        safe {
            $path = selectFolder "検索する TSV インデックスのフォルダを選んでください" ${indexDir} $script:indexDialog.Window
            if ($path) {
                addIndexDialogFolder $path
            }
        }
    })
    $dialog.FindName("RemoveButton").Add_Click({ safe { removeIndexDialogFolder } })
    $dialog.FindName("UpButton").Add_Click({ safe { moveIndexDialogFolder -1 } })
    $dialog.FindName("DownButton").Add_Click({ safe { moveIndexDialogFolder 1 } })
    $dialog.FindName("OkButton").Add_Click({ $script:indexDialog.Window.DialogResult = $true })
    $list.Add_PreviewDragOver({ onFolderDragOver @args })
    $list.Add_PreviewDrop({
        param ($sender, $e)
        safe {
            foreach ($folder in (getDroppedFolders $e)) {
                addIndexDialogFolder $folder
            }
        }
        $e.Handled = $true
    })
    $list.Add_KeyDown({
        param ($sender, $e)
        if ($e.Key -eq "Delete") {
            safe { removeIndexDialogFolder }
        }
    })

    if ($dialog.ShowDialog()) {
        writeIndexFolders @($list.Items | ForEach-Object { [string]$_.Content })
        $script:indexSummary = $null
        loadIndexTree
        refreshIndexSummary
        setStatus "検索対象インデックスを保存しました"
    }
    $script:indexDialog = $null
}
# ---- イベント ----

$ui.WordBox.AddHandler([System.Windows.Controls.Primitives.TextBoxBase]::TextChangedEvent, [System.Windows.Controls.TextChangedEventHandler] { safe { updateWordNotice } })
$ui.WordBox.Add_PreviewKeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Return") {
        safe {
            if (!$script:search) {
                startSearch
            }
        }
        $e.Handled = $true
    }
})
$ui.SearchButton.Add_Click({ safe { startSearch } })
$ui.RegexCheck.Add_Click({
    safe {
        writeSearchOption @{ UseRegex = [bool]$ui.RegexCheck.IsChecked }
        updateWordNotice
    }
})
$ui.CaseCheck.Add_Click({ safe { writeSearchOption @{ CaseSensitive = [bool]$ui.CaseCheck.IsChecked } } })
$ui.FileFilterBox.Add_TextChanged({
    $ui.FileFilterPlaceholder.Visibility = if ($ui.FileFilterBox.Text -eq "") { "Visible" } else { "Collapsed" }
})
# 対象ファイルは入力を終えたとき（フォーカスが外れたとき・検索したとき）に保存する
$ui.FileFilterBox.Add_LostFocus({ safe { writeSearchOption @{ FileFilter = $ui.FileFilterBox.Text.Trim() } } })
$ui.FileFilterBox.Add_KeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Return") {
        safe {
            if (!$script:search) {
                startSearch
            }
        }
        $e.Handled = $true
    }
})
$ui.ChangeIndexButton.Add_Click({ safe { showIndexDialog } })
# ツリーのチェックボックスのクリック（チェックの状態は IndexNode.IsChecked に反映済み）
$ui.IndexTree.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent, [System.Windows.RoutedEventHandler] { safe { onIndexTreeChecked } })
$ui.IndexTree.Add_PreviewKeyDown({
    param ($sender, $e)
    # スペースで選択中のフォルダのチェックを切り替える
    $node = $ui.IndexTree.SelectedItem
    if ($e.Key -eq "Space" -and $node -and !$node.IsPlaceholder) {
        safe {
            $node.Toggle()
            onIndexTreeChecked
        }
        $e.Handled = $true
    }
})
$ui.CheckAllIndexButton.Add_Click({ safe { setAllIndexChecked $true } })
$ui.UncheckAllIndexButton.Add_Click({ safe { setAllIndexChecked $false } })
$ui.GoIndexTabButton.Add_Click({ $ui.Tabs.SelectedItem = $ui.IndexTab })
$ui.FilterBox.Add_TextChanged({
    $ui.FilterPlaceholder.Visibility = if ($ui.FilterBox.Text -eq "") { "Visible" } else { "Collapsed" }
    $script:filterTimer.Stop()
    $script:filterTimer.Start()
})
$ui.ResultGrid.Add_SelectionChanged({ safe { showDetail } })
$ui.ResultGrid.Add_MouseDoubleClick({
    param ($sender, $e)
    # 行の上でのダブルクリックだけを対象にする（列見出し・スクロールバーは除く）
    $element = $e.OriginalSource
    while ($element -and !($element -is [System.Windows.Controls.DataGridRow])) {
        if ($element -is [System.Windows.Controls.Primitives.DataGridColumnHeader] -or $element -is [System.Windows.Controls.Primitives.ScrollBar]) {
            return
        }
        $element = [System.Windows.Media.VisualTreeHelper]::GetParent($element)
    }
    if ($element) {
        safe { openSource }
    }
})
$ui.ResultGrid.Add_PreviewKeyDown({
    param ($sender, $e)
    if ($e.Key -eq "Return") {
        safe { openSource }
        $e.Handled = $true
    } elseif ($e.Key -eq "C" -and [System.Windows.Input.Keyboard]::Modifiers -eq "Control") {
        safe { copySelectedRows }
        $e.Handled = $true
    }
})
$ui.MenuOpen.Add_Click({ safe { openSource ${openModeNormal} } })
$ui.MenuOpenReadOnly.Add_Click({ safe { openSource ${openModeReadOnly} } })
$ui.MenuOpenNew.Add_Click({ safe { openSource ${openModeNew} } })
$ui.OpenModeCombo.Add_SelectionChanged({
    safe {
        # 起動時の読み込みでは保存しない（設定していない利用者の setting.config を作らないため）
        if (-not $script:loadingOpenMode) {
            writeOpenMode (getOpenMode)
        }
        updateOpenMenu
    }
})
$ui.MenuOpenFolder.Add_Click({ safe { openSourceFolder } })
$ui.OpenButton.Add_Click({ safe { openSource } })
$ui.OpenFolderButton.Add_Click({ safe { openSourceFolder } })
# プレビューのセルをクリックすると、その値をコピーできるように選ぶ（Shift＋クリック・ドラッグで範囲、Ctrl+C でコピー）
$script:previewTable = $null
$ui.PreviewRows.Add_PreviewMouseLeftButtonDown({
    param ($sender, $e)
    safe {
        $cell = getPreviewCell $e.OriginalSource
        if ($cell -and $script:previewTable) {
            $extend = [System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift
            $script:previewTable.Select($cell, [bool]$extend)
            [void]$ui.PreviewScroll.Focus()
        }
    }
})
$ui.PreviewRows.Add_MouseMove({
    param ($sender, $e)
    if ($e.LeftButton -ne "Pressed") {
        return
    }
    safe {
        $cell = getPreviewCell $e.OriginalSource
        if ($cell -and $script:previewTable) {
            $script:previewTable.Select($cell, $true)
        }
    }
})
$ui.PreviewRows.Add_PreviewMouseRightButtonDown({
    param ($sender, $e)
    safe {
        # 右クリックしたセルが選ばれていなければ、そのセルだけを選ぶ
        $cell = getPreviewCell $e.OriginalSource
        if ($cell -and $script:previewTable -and !$cell.IsSelected) {
            $script:previewTable.Select($cell, $false)
        }
    }
})
$ui.PreviewScroll.Add_PreviewKeyDown({
    param ($sender, $e)
    if ($e.Key -eq "C" -and [System.Windows.Input.Keyboard]::Modifiers -eq "Control") {
        safe { copyPreviewSelection }
        $e.Handled = $true
    }
})
$ui.MenuPreviewCopy.Add_Click({ safe { copyPreviewSelection } })
$ui.MenuPreviewCopyRow.Add_Click({
    safe {
        if ($script:previewTable -and $script:previewTable.HasSelection) {
            # 選んでいるセルのある行をすべて選んでからコピーする
            $selected = $null
            foreach ($row in $script:previewTable.Rows) {
                $selected = @($row.Cells | Where-Object { $_.IsSelected })[0]
                if ($selected) {
                    break
                }
            }
            if ($selected) {
                $script:previewTable.SelectRow($selected)
            }
        }
        copyPreviewSelection
    }
})
# プレビューの列見出しの右端をドラッグすると、その列（PreviewColumn）の幅が変わる（各行のセルも同じ列を参照しているため一緒に変わる）
$ui.PreviewHeader.AddHandler(
    [System.Windows.Controls.Primitives.Thumb]::DragDeltaEvent,
    [System.Windows.Controls.Primitives.DragDeltaEventHandler] {
        param ($sender, $e)
        safe {
            $column = $e.OriginalSource.DataContext
            if ($column -is [WinGrep.PreviewColumn]) {
                $column.Width = $column.Width + $e.HorizontalChange
            }
        }
    })
$ui.MenuCopy.Add_Click({ safe { copySelectedRows } })
$ui.MenuCopyPath.Add_Click({ safe { copySourcePath } })
$ui.ExportButton.Add_Click({ safe { exportResults } })

# ============================================================================
# ［9 Office 強制終了］
# ============================================================================

$script:processes = @()

function refreshProcesses {
    $selectedIds = @($ui.ProcessGrid.SelectedItems | ForEach-Object { $_.Id })
    $script:processes = @(getOfficeProcesses | Sort-Object @{ Expression = { !$_.Background } }, @{ Expression = { $_.StartTime } })

    $rows = New-Object System.Collections.ArrayList
    foreach ($process in $script:processes) {
        $row = New-Object WinGrep.ProcRow
        $row.Id = $process.Id
        $row.AppName = $process.AppName
        $row.Background = $process.Background
        $row.StateText = if ($process.Background) { "⚠ バックグラウンド" } else { "画面に表示中" }
        $row.StartText = formatTime $process.StartTime
        $row.MemoryText = "$($process.MemoryMB.ToString('N0')) MB"
        $row.TitleText = if ($process.Title) { $process.Title } else { "（なし）" }
        [void]$rows.Add($row)
    }
    $ui.ProcessGrid.ItemsSource = $rows
    foreach ($row in $rows) {
        if ($selectedIds -contains $row.Id) {
            [void]$ui.ProcessGrid.SelectedItems.Add($row)
        }
    }

    $background = @($script:processes | Where-Object { $_.Background }).Count
    $visible = $script:processes.Count - $background
    if ($script:processes.Count -eq 0) {
        $ui.ProcessSummaryText.Text = "実行中の Excel・Word・PowerPoint はありません。（$(Get-Date -Format 'H:mm:ss') 時点）"
    } else {
        $ui.ProcessSummaryText.Text = "$($script:processes.Count) 件（バックグラウンド $background 件・画面に表示中 $visible 件） ・ $(Get-Date -Format 'H:mm:ss') 時点"
    }
    $ui.KillBackgroundButton.Content = "バックグラウンドのみ終了（$background 件）"
    $ui.KillBackgroundButton.IsEnabled = $background -gt 0
    $ui.KillAllButton.IsEnabled = $script:processes.Count -gt 0
    $ui.KillSelectedButton.IsEnabled = $ui.ProcessGrid.SelectedItems.Count -gt 0
    updateKillBadge $background
}

function updateKillBadge {
    param (
        $background = $null
    )

    if ($null -eq $background) {
        $background = @(getOfficeProcesses | Where-Object { $_.Background }).Count
    }
    # 変換中はバックグラウンドの Excel 等があって当然なので、印を付けない
    $ui.KillTabHeader.Text = if ($background -gt 0 -and !(isConverting)) { "⚠ 9 Office 強制終了" } else { "9 Office 強制終了" }
}

function killProcesses {
    param (
        [object[]]$targets
    )

    if ($targets.Count -eq 0) {
        return
    }

    $describe = {
        param ([object[]]$list)
        @(${officeProcessNames}.Values | ForEach-Object {
            $name = $_
            $count = @($list | Where-Object { $_.AppName -eq $name }).Count
            if ($count -gt 0) { "$name $count 件" }
        }) -join "・"
    }
    $visibleTargets = @($targets | Where-Object { !$_.Background })
    $message = if ($visibleTargets.Count -gt 0) {
        "画面に表示中の $(& $describe $visibleTargets) を含む $($targets.Count) 件を、保存せずに終了します。`n保存していない内容は失われます。よろしいですか？"
    } else {
        "バックグラウンドの $(& $describe $targets) を終了します。よろしいですか？"
    }
    $default = if ($visibleTargets.Count -gt 0) { "No" } else { "Yes" }
    if (isConverting) {
        $message = "変換中です。バックグラウンドのプロセスを終了すると、変換中のファイルは失敗扱いになります。`n`n" + $message
        $default = "No"
    }
    if ((showMessage $message "YesNo" "Warning" $default) -ne "Yes") {
        return
    }

    $results = @(stopOfficeProcesses @($targets | ForEach-Object { $_.Id }))
    $stopped = @($results | Where-Object { $_.Stopped }).Count
    $failures = @($results | Where-Object { !$_.Stopped } | ForEach-Object { "PID $($_.Id) を終了できませんでした：$($_.Message)" })
    setStatus (@("$stopped 個のプロセスを終了しました。") + $failures -join "　")
    Start-Sleep -Milliseconds 300
    refreshProcesses
}

$script:processTimer = newTimer 5000 { safe { refreshProcesses } }

$ui.RefreshProcessButton.Add_Click({ safe { refreshProcesses } })
$ui.ProcessGrid.Add_SelectionChanged({ $ui.KillSelectedButton.IsEnabled = $ui.ProcessGrid.SelectedItems.Count -gt 0 })
$ui.KillBackgroundButton.Add_Click({ safe { refreshProcesses; killProcesses @($script:processes | Where-Object { $_.Background }) } })
$ui.KillSelectedButton.Add_Click({
    safe {
        $ids = @($ui.ProcessGrid.SelectedItems | ForEach-Object { $_.Id })
        killProcesses @($script:processes | Where-Object { $ids -contains $_.Id })
    }
})
$ui.KillAllButton.Add_Click({ safe { refreshProcesses; killProcesses $script:processes } })

# ============================================================================
# ウィンドウ全体
# ============================================================================

$ui.CloseButton.Add_Click({ $window.Close() })

$ui.Tabs.Add_SelectionChanged({
    param ($sender, $e)
    # 中の表・一覧の選択変更も伝わってくるため、タブの切り替えだけを扱う
    if ($e.OriginalSource -ne $ui.Tabs) {
        return
    }
    safe {
        if ($ui.Tabs.SelectedItem -eq $ui.KillTab) {
            refreshProcesses
            $script:processTimer.Start()
        } else {
            $script:processTimer.Stop()
        }
        if ($ui.Tabs.SelectedItem -eq $ui.IndexTab) {
            refreshConversionState
        }
    }
})

$window.Add_Activated({
    safe {
        # 変換対象フォルダがほかの画面で変更されていれば読み直す
        if ((getTargetsKey @(getTargetFolders)) -ne $script:savedTargets) {
            loadTargets
            setStatus "変換対象フォルダがほかで変更されたため、読み直しました"
        }
        foreach ($item in $script:targetItems) {
            updateFolderItemStatus $item
        }
        if (!(isConverting)) {
            refreshConversionState
        }
        updateSearchTarget
        if ($ui.Tabs.SelectedItem -eq $ui.KillTab) {
            refreshProcesses
        } else {
            updateKillBadge
        }
    }
})

$window.Add_PreviewKeyDown({
    param ($sender, $e)
    $modifiers = [System.Windows.Input.Keyboard]::Modifiers
    if ($e.Key -eq "F" -and $modifiers -eq "Control") {
        $ui.Tabs.SelectedItem = $ui.SearchTab
        $ui.WordBox.Focus() | Out-Null
        $textBox = $ui.WordBox.Template.FindName("PART_EditableTextBox", $ui.WordBox)
        if ($textBox) {
            $textBox.SelectAll()
        }
        $e.Handled = $true
    } elseif ($e.Key -eq "F" -and $modifiers -eq ([System.Windows.Input.ModifierKeys]::Control -bor [System.Windows.Input.ModifierKeys]::Shift)) {
        $ui.Tabs.SelectedItem = $ui.SearchTab
        $ui.FilterBox.Focus() | Out-Null
        $e.Handled = $true
    } elseif ($e.Key -eq "F5") {
        safe {
            if ($ui.Tabs.SelectedItem -eq $ui.KillTab) {
                refreshProcesses
            } else {
                refreshConversionState
                refreshIndexSummary
                loadIndexTree
            }
        }
        $e.Handled = $true
    } elseif ($e.Key -eq "Escape" -and $script:search) {
        cancelSearch
        $e.Handled = $true
    }
})

$window.Add_Closing({
    param ($sender, $e)
    # 変換はウィンドウを出さずに動いているため、閉じる前にどうするか聞く
    if (isConverting) {
        $answer = showMessage ("変換中です。`n`n［はい］変換を中止してから閉じる（変換中のファイルが終わったところで止まります）`n" +
            "［いいえ］変換を続けたまま閉じる（もう一度開くと進み具合を表示します）`n［キャンセル］閉じない") "YesNoCancel" "Question" "Cancel"
        if ($answer -eq "Cancel") {
            $e.Cancel = $true
            return
        }
        if ($answer -eq "Yes") {
            [System.IO.File]::WriteAllText(${stopRequestFile}, "", ${utf8Bom})
        }
    }
    if ($script:search) {
        $script:search.Shared.Stop = $true
    }
})

$window.Add_Loaded({
    safe {
        if ($ui.Tabs.SelectedItem -eq $ui.SearchTab) {
            $ui.WordBox.Focus() | Out-Null
        }
    }
})

# ---- 起動 ----

loadTargets
setSearchOptionToUi (readSearchOption)
setOpenMode (readOpenMode)
updateOpenMenu
loadHistory
refreshConversionState
loadIndexTree
updateWordNotice
updateKillBadge
refreshIndexSummary

# 前回の画面で起動した変換が続いていれば、進み具合を表示する
$runningConversion = findRunningConversion
if ($runningConversion) {
    adoptConversion $runningConversion
}

# 起動時のタブ：変換中・中断中、またはインデックスが無ければ［1 インデックス作成］、それ以外は［2 検索］
$openIndexTab = $runningConversion -or ($script:conversionState -and $script:conversionState.Pending -gt 0) -or !(testIndexExists)
$ui.Tabs.SelectedItem = if ($openIndexTab) { $ui.IndexTab } else { $ui.SearchTab }
setStatus ""

try {
    [void]$window.ShowDialog()
} finally {
    if ($script:search) {
        $script:search.PS.Stop()
    }
    $mutex.ReleaseMutex()
    $mutex.Dispose()
}
