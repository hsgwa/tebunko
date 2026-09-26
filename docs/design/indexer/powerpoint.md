# PowerPoint

Word・PowerPoint で共通の処理（抽出の流れ・起動と終了・ファイルの読み取り）は [Word・PowerPoint の共通処理と Office アプリの管理](office-apps.md) を参照。

PowerPoint の旧形式（`.ppt` 等）の変換（[PowerPoint の旧形式の変換](#powerpoint-の旧形式の変換extractwithpowerpoint)）、スライド・ノート・図形・コメントごとの読み取り（[PowerPoint のテキスト読み取りと TSV の場所](#powerpoint-のテキスト読み取りと-tsv-の場所readpptxunits)）、PowerPoint に固有の注意点（[PowerPoint の注意点・既知の問題](known-issues.md#powerpoint-の注意点既知の問題)）を扱う。

## PowerPoint の旧形式の変換（`extractWithPowerPoint`）

[Word・PowerPoint の抽出処理](office-apps.md#wordpowerpoint-の抽出処理extractdocument)で ZIP ではないと判定したファイル（`.ppt`、パスワード付き、拡張子と中身が異なるもの）を、PowerPoint で `.pptx` に変換する。

| 項目 | 仕様 |
|---|---|
| 開く前の確認 | ファイルの先頭が複合ドキュメント形式（旧形式・パスワード付きの Office ファイル。`D0 CF 11 E0 A1 B1 1A E1`、`isCompoundFile`）でなければ、PowerPoint で開かずに失敗とする（`ファイルが壊れているか、PowerPointのファイルではありません（新形式（ZIP）でも旧形式でもない内容です）。`）。PowerPoint はテキストなどのファイルもアウトラインとして開いてしまい、文字化けした内容になるため（◎ テストデータの `PowerPoint\異常系\壊れたファイル.pptx`）。Word はテキスト・HTML・RTF も正しく読めるため、この確認はしない |
| 開き方 | `Presentations.Open("パス::dummy::", ReadOnly=True, Untitled=False, WithWindow=False)` |
| パスワード付きファイル | ファイル名末尾の `::<パスワード>::` により、ダイアログを出さずに例外 |
| 保存 | `SaveAs(converted.pptx, 24 = ppSaveAsOpenXMLPresentation)` |
| アプリの設定 | `DisplayAlerts = 1`（ppAlertsNone）、`AutomationSecurity = 3`。PowerPoint はウィンドウを隠せないため、ファイルをウィンドウ無しで開く |

## PowerPoint のテキスト読み取りと TSV の場所（`readPptxUnits`）

`.pptx` は [Word・PowerPoint のテキスト読み取り](office-apps.md#wordpowerpoint-のテキスト読み取りscriptssharedofficeoffice_readerps1)の共通の読み取り（`readXmlLines`、DrawingML の名前空間 `a:`）で読み、次の場所ごとに TSV にする。

| 場所 | 読み取り元 |
|---|---|
| `スライドNNN` | `ppt/presentation.xml` の `p:sldIdLst` の順（= 表示順）に、リレーションシップからスライドの XML を引く。ファイル名の番号（`slide3.xml` 等）は表示順と一致しないことがあるため使わない |
| `スライドNNN（非表示）` | 非表示スライド（`<p:sld show="0">`。`show="false"` も同じ） |
| `スライドNNN_ノート` | スライドのリレーションシップ（種類 `notesSlide`）が指すノートの XML |
| `ヘッダー・フッター` | 各スライドのフッターのプレースホルダー（`p:ph` の `type` が `ftr`）。各スライドに同じ内容が並ぶため、全スライド分をまとめて重複する行は除く（Word の `ヘッダー・フッター` と同じ扱い） |
| `スライドNNN[図形]` | SmartArt（`dgm:relIds` の `r:dm` が指す `ppt/diagrams/dataN.xml`）とグラフ（`c:chart` の `r:id` が指す `ppt/charts/chartN.xml` のタイトル・軸ラベル・系列名・項目名。数値は読まない）。図形 1 つを 1 行にし、段落はスペースでつなぐ |
| `スライドNNN[コメント]` | スライドのリレーションシップ（種類が `comments` で終わるもの）が指すコメント。旧形式（`ppt/comments/commentN.xml` の `p:cm` の `p:text`）と新形式（`ppt/comments/modernComment_*.xml` の `p188:cm` の本文と返信）。コメントの後に返信を 1 件ずつ並べる。作成者名は読まない |

- 図形・グループ内の図形・表のテキストを読む。スライド番号・日付・ヘッダー・フッター・スライド画像のプレースホルダー（`p:ph` の `type` が `sldNum` `dt` `hdr` `ftr` `sldImg`）は、スライドの場所には含めない（フッターは上の `ヘッダー・フッター` にまとめる）。
- ノートのプレースホルダーのうち、ヘッダー・フッター・日付・ページ番号・スライド画像は読まない。
- テキストボックス・図形の文字はスライドの本文にする（`[図形]` にしない）。スライドの文字はほとんどが図形（プレースホルダーを含む）で、分けると本文が空になるため。
- 非表示スライドの図形・コメントは `スライドNNN（非表示）[図形]` のように、非表示と分かる名前にする。
- スライドマスター・レイアウトのテキスト、埋め込みオブジェクトは読まない。

TSV の例（配置・命名の規則は [配置・命名規則](index-format.md#配置命名規則) の [配置・命名規則](index-format.md#配置命名規則)。クロール対象フォルダが `C:\data\営業`（インデックス名 `営業`）の場合）:

| 元ファイル（クロール対象フォルダからの相対パス） | 場所 | 出力 TSV |
|---|---|---|
| `提案.pptx` | 3 枚目のスライド | `work/index/営業/提案.pptx/スライド003.tsv` |
| `提案.pptx` | 3 枚目（非表示スライド） | `work/index/営業/提案.pptx/スライド003（非表示）.tsv` |
| `提案.pptx` | 3 枚目の発表者ノート | `work/index/営業/提案.pptx/スライド003%5Fノート.tsv` |
| `提案.pptx` | スライドのフッター | `work/index/営業/提案.pptx/ヘッダー・フッター.tsv` |

- スライドの番号は 3 桁以上の 0 埋め（ファイル名の順 = スライドの表示順）。
