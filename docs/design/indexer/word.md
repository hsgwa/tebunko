# Word

Word・PowerPoint で共通の処理（抽出の流れ・起動と終了・ファイルの読み取り）は [Word・PowerPoint の共通処理と Office アプリの管理](office-apps.md) を参照。

Word 文書の旧形式（`.doc` 等）の変換（[Word の旧形式の変換](#word-の旧形式の変換extractwithword)）、本文のページ・ヘッダー/フッター・脚注・図形・コメントごとの読み取り（[Word のテキスト読み取りと TSV の場所](#word-のテキスト読み取りと-tsv-の場所readdocxunits)）、Word に固有の注意点（[Word の注意点・既知の問題](known-issues.md#word-の注意点既知の問題)）を扱う。

## Word の旧形式の変換（`extractWithWord`）

[Word・PowerPoint の抽出処理](office-apps.md#wordpowerpoint-の抽出処理extractdocument)で ZIP ではないと判定したファイル（`.doc`、パスワード付き、拡張子と中身が異なるもの）を、Word で `.docx` に変換する。

| 項目 | 仕様 |
|---|---|
| 開き方 | `Documents.Open(パス, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False, パスワード類="dummy", ..., Visible=False)` |
| パスワード付きファイル | ダミーのパスワードにより、ダイアログを出さずに「パスワードが正しくありません」の例外 |
| 保存 | `Repaginate()` でページ割りを確定させてから `SaveAs2(converted.docx, 12 = wdFormatXMLDocument)` |
| アプリの設定 | `Visible = False`、`DisplayAlerts = 0`（wdAlertsNone）、`AutomationSecurity = 3` |

- `Repaginate()` を呼ぶのは、保存時に記録されるページ区切り（6.5 のページの目安）を確定させるため。呼ばないと、同じ内容の文書でも変換のたびにページ区切りの位置が変わることがある（◎）。

## Word のテキスト読み取りと TSV の場所（`readDocxUnits`）

`.docx` は [Word・PowerPoint のテキスト読み取り](office-apps.md#wordpowerpoint-のテキスト読み取りscriptssharedofficeoffice_readerps1)の共通の読み取り（`readXmlLines`、WordprocessingML の名前空間 `w:`）で読み、次の場所ごとに TSV にする。

| 場所 | 読み取り元 |
|---|---|
| `ページNNN` | `word/document.xml`（本文） |
| `ヘッダー・フッター` | `word/header*.xml` → `word/footer*.xml` の順。セクションごとに同じ内容が並ぶため、重複する行は除く |
| `脚注` | `word/footnotes.xml`・`word/endnotes.xml`（区切り線は文字が無いため出力されない） |
| `ページNNN[図形]` | 本文のテキストボックス・図形内の文字（`w:txbxContent`）、SmartArt（`dgm:relIds` の `r:dm` が指す `word/diagrams/dataN.xml`）、グラフ（`c:chart` の `r:id` が指す `word/charts/chartN.xml` のタイトル・軸ラベル・系列名・項目名。数値は読まない）。図形 1 つを 1 行にし、段落はスペースでつなぐ。ページは図形を置いた段落のページ |
| `ページNNN[コメント]` | `word/comments.xml` のコメント（返信も 1 件ずつ）。ページは本文の `w:commentReference` の位置。本文に参照の無いコメント（ヘッダー・脚注に付けたもの等）は `文書[コメント]` にまとめる。作成者名は読まない |

- 図形・コメントの場所の決まりは [配置・命名規則](index-format.md#配置命名規則)の「図形・コメントの場所」。SmartArt は描画用の `diagrams/drawingN.xml` に同じ文字があるが、重複するため読まない。
- ヘッダー・フッター・脚注の中のテキストボックスは、`[図形]` に分けずにその場所の行にする。
- **ページ番号は目安**。ファイルにはページの情報が無いため、以下で数える。
  - Word が保存時に記録したページ区切り（`w:lastRenderedPageBreak`）。本文にこれが 1 つも無い（Word 以外で作られた）文書では使わない。
  - 手動の改ページ（`w:br w:type="page"`）・段落前で改ページ（`w:pageBreakBefore`）・セクション区切り（種類が「連続」以外）。これらの直後、本文が出る前にある `w:lastRenderedPageBreak` は同じ区切りなので数えない（Word は手動の改ページの後に `w:lastRenderedPageBreak` を記録する場合としない場合がある ◎）。
  - 段落・表の行は、最初の文字があるページに属する。

TSV の例（配置・命名の規則は [配置・命名規則](index-format.md#配置命名規則) の [配置・命名規則](index-format.md#配置命名規則)。クロール対象フォルダが `C:\data\営業`（インデックス名 `営業`）の場合）:

| 元ファイル（クロール対象フォルダからの相対パス） | 場所 | 出力 TSV |
|---|---|---|
| `報告書.docx` | 2 ページ目 | `work/index/営業/報告書.docx/ページ002.tsv` |
| `報告書.docx` | ヘッダー・フッター | `work/index/営業/報告書.docx/ヘッダー・フッター.tsv` |
| `報告書.docx` | 脚注・文末脚注 | `work/index/営業/報告書.docx/脚注.tsv` |

- ページ番号は 3 桁以上の 0 埋め（ファイル名の順 = ページの順）。
