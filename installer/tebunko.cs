// tebunko.exe（インストーラー版の起動口）
//
// インストーラー（installer\tebunko.iss）が、scripts\ と並べて入れる小さな exe。zip 版には入れない（zip 版の起動口は tebunko.bat）。
// していることは tebunko.bat と同じで、同じフォルダの scripts\tebunko\gui.ps1 を、このプロセスだけの実行ポリシー RemoteSigned で起動する。
// bat との違い:
//   ・窓のアプリ（/target:winexe）で、PowerShell も窓を作らずに起動する（CreateNoWindow）。bat のように conhost.exe を通す必要が無い
//   ・Mark-of-the-Web は消さない。インストーラーが書いたファイルには付かないため
//   ・PowerShell が 0 以外で終わり、エラー出力があれば、その内容をメッセージで出す（実行ポリシー・制限言語モードで
//     gui.ps1 が読み込めないときは、gui.ps1 の trap に届かず何も出ないため）
//
// エラー出力を受け取るため、画面が閉じるまでこのプロセスも残る（窓は無い）。先に終わると、PowerShell が書き込めずに止まる。
// 起動中はミューテックス tebunko-installed-app を持ち、インストーラーはこれを見て「tebunko を閉じてください」と出す（AppMutex）。
//
// ビルドは tools\new_installer.ps1 が Windows 標準の .NET Framework の csc.exe で行う（C# 5 まで。道具を足さない）。
// 版の情報（AssemblyVersion など）は、ビルドのときに別のファイルで足す。
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

static class Program
{
    const string AppTitle = "tebunko";
    const string MutexName = "tebunko-installed-app";
    // メッセージに出すエラー出力の長さの上限
    const int MaxErrorLength = 2000;

    [STAThread]
    static int Main()
    {
        string root = AppDomain.CurrentDomain.BaseDirectory;
        string gui = Path.Combine(root, @"scripts\tebunko\gui.ps1");
        if (!File.Exists(gui))
        {
            ShowError("tebunko の本体が見つかりません。インストールし直してください。\n\n" + gui);
            return 1;
        }
        // PATH から探さず、Windows に入っている PowerShell 5.1 を指す
        string powershell = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System),
            @"WindowsPowerShell\v1.0\powershell.exe");

        using (var mutex = new Mutex(false, MutexName))
        {
            var info = new ProcessStartInfo(powershell, BuildArguments(gui));
            info.UseShellExecute = false;
            info.CreateNoWindow = true;
            info.WorkingDirectory = root;
            info.RedirectStandardError = true;
            // PowerShell は、窓の無いコンソールのコードページ（OEM）でエラー出力を書く
            info.StandardErrorEncoding = Encoding.GetEncoding(CultureInfo.CurrentCulture.TextInfo.OEMCodePage);

            var errors = new StringBuilder();
            Process process;
            try
            {
                process = Process.Start(info);
            }
            catch (Exception e)
            {
                ShowError("PowerShell を起動できませんでした。\n\n" + e.Message);
                return 1;
            }
            process.ErrorDataReceived += (sender, e) =>
            {
                if (e.Data != null)
                {
                    lock (errors) { errors.AppendLine(e.Data); }
                }
            };
            process.BeginErrorReadLine();
            process.WaitForExit();

            int exitCode = process.ExitCode;
            string text;
            lock (errors) { text = errors.ToString().Trim(); }
            if (exitCode != 0 && text.Length > 0)
            {
                ShowError("tebunko を起動できませんでした。\n\n" + Truncate(text));
            }
            return exitCode;
        }
    }

    // gui.ps1 を起動する PowerShell の引数（tebunko.bat と同じ。-WindowStyle Hidden は窓を作らないので要らない）
    internal static string BuildArguments(string gui)
    {
        return "-NoProfile -STA -ExecutionPolicy RemoteSigned -File \"" + gui + "\"";
    }

    static string Truncate(string text)
    {
        return text.Length <= MaxErrorLength ? text : text.Substring(0, MaxErrorLength) + "…";
    }

    static void ShowError(string message)
    {
        MessageBox.Show(message, AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
