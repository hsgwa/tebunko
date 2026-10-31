# 部品ごとの関数（TSV と本文インデックス）

扱うこと: TSV の名前と場所・整形（`fs.ps1`・`index_name.ps1`・`text.ps1`・`tsv_file.ps1`）、インデックスへの配置（`index_store.ps1`）、本文インデックスの形式と読み書き（`content_index_format.ps1`・`content_index_store.ps1`）の関数一覧。扱わないこと: 検索そのものの関数（[部品ごとの関数（検索・スレッド・元のファイル・画面）](search.md)）。先に読むページ: [部品から関数一覧を引く](index.md)。

## TSV の名前と場所（`shared/core/fs.ps1`・`tebunko/index/index_name.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `encodeIndexPlace` | place | string | TSV のファイル名に入れる場所を符号化する（ファイル名禁止文字・制御文字・`_`・`%` を `%XX` に） | [インデックスのファイルの形](../index-data/format.md#場所の符号化encodeindexplace--decodeindexplace) | toIndexFileName |
| `decodeIndexPlace` | place | string | `encodeIndexPlace` の `%XX` を元に戻す（それ以外の `%` はそのまま） | 同上 | convertIndexFileNameToPlace |
| `convertPlaceBaseToFixedFileName` | base | string（固定名 / 空） | 場所の元の名前（`ページ001` など）が決まった種類なら、固定の英語のファイル名（`page_001` など）を返す。決まった種類でなければ空文字列 | [インデックスのファイルの形](../index-data/format.md#配置命名規則) | toIndexFileName |
| `convertFixedFileNameToPlaceBase` | fileBase | string（元の名前 / 空） | `convertPlaceBaseToFixedFileName` の逆（`page_001` → `ページ001`）。固定名でなければ空文字列 | 同上 | convertIndexFileNameToPlace |
| `toIndexFileName` | place | string | インデックスの TSV のファイル名 `<場所>.tsv`（場所は `convertPlaceBaseToFixedFileName` の固定名、無ければ `encodeIndexPlace`）。図形・コメント・ヘッダー・フッターは、末尾に英語の種類 `[shape]` `[comment]` `[header_footer]` を付ける。元のファイル名はフォルダ名にするため入れない。`$maxFileNameLength`（255）文字を超えれば例外 | [インデックスのファイルの形](../index-data/format.md#配置命名規則) | インデックス作成（Excel・Word・PowerPoint） |
| `convertIndexFileNameToPlace` | fileNameWithoutExtension | string | `toIndexFileName` の逆。図形・コメント・ヘッダー・フッターの種類（`[shape]` `[comment]` `[header_footer]`）を外してから、固定名なら `convertFixedFileNameToPlaceBase`、そうでなければ `decodeIndexPlace` で場所の元の名前に戻す | 同上 | getIndexFolderBooks |
| `splitObjectPlace` | place | `@{Base; Kind}` | 図形・コメント・ヘッダー・フッターの場所（`<元の場所>[図形]` 等）を、元の場所と種類に分ける。ふつうの場所は Kind が空 | [インデックスのファイルの形](../index-data/format.md#配置命名規則)「図形・コメントの場所」 | 元のファイルを開く、convertPlaceToContentIndexMeta |
| `newObjectPlaceKindPattern` | kinds（省略は全部） | string | 場所の末尾の種類（`[図形]` `[埋め込み<N>]` など）だけを当てる正規表現の文字列を、`placeKindFileNames` の表から組み立てる。検索から外す場所（`newPlaceExclude`）と画面の判定（`HitRow.ObjectPlaceRegex`）が使う | [インデックスのファイルの形](../index-data/format.md#配置命名規則)「図形・コメントの場所」 | newPlaceExclude、画面（HitRow） |
| `describePlace` | book, place | `@{Place; Kind}` | 場所ごとの表記と種別（見出しの要約・検索結果ファイル・コピーに出す文字。`[シート]売上`・図形、`3 ページ（目安）`・本文 など） | [出力フォーマット](../search/output.md#出力フォーマットexcel-ブックcsvtsv) | 画面・`toResultLine` |
| `describeHitPlace` | place, isExcel, isObjectPlace, matchCell, matchCount, lineNumber, isText（既定 `$false`） | string | 結果の表の「場所」列・プレビューの題に出す、行ごとの表記（Excel は場所ごとの表記にセル番地を足す。`[シート]売上!B12`・`[シート]売上!B12 ほか 2`・`[シート]売上 12 行目`。図形・コメントは `[シート]売上!D5`。ヘッダー・フッターはセル番地が無く `[シート]売上`。Word・PowerPoint は場所ごとのまま。テキストは `isText` が真なら行番号だけ `12 行目`） | [結果の表](../gui/search-tab.md#結果の表) | 画面（`prepareHitRow`） |
| `toLongPath` | path | string | ファイル操作に渡すパスの先頭に `\\?\`（ネットワークのパスは `\\?\UNC\`）を付け、260 文字を超えるパスも扱えるようにする。付いていればそのまま | [入れ替えと書き出し](../index-data/publish.md#長いパス260-文字超の扱い) | インデックス作成・検索 |
| `fromLongPath` | path | string | `toLongPath` で付けた `\\?\` を外す（`Get-ChildItem` の `FullName` から相対パスを求めるため） | 同上 | インデックス作成・検索 |
| `testPathUnder` | path, root | bool | path が root の下（root そのものは含まない）にあるか。`..`・`.`・最後の `.` や空白・`\\?\` を解決してから比べる。状態のファイルから読んだ名前でフォルダを消す前に使う | [ファイルの読み書きの範囲](../../safety/file-access.md) | インデックス・作業フォルダの削除, `IndexCatalog` |
| `trimTrailingSeparator` | path | string | パスの末尾の `\`・`/` を落とす（ドライブのルート・`\\?\` だけのパスは落とさない）。リンクの判定・削除の入口で使う | 同上 | `testPathIsLink`・`removeTreeNoFollow` |
| `removeDirectoryRetry` | path, tries（既定 3）, waitMilliseconds（既定 200） | – | フォルダを中身ごと削除する（`removeTreeNoFollow`。中のリンクは辿らない）。ほかのアプリが一時的に掴んでいることがあるため、少し待って数回試す。消したあとに `[System.IO.Directory]::Exists` でフォルダが残っていれば、成功にせず例外にする（呼ぶ側は、消えたことが確かめられてから記録する）。試すたびに `Exists` が偽なら `getPathState` で見る。無いと確かめられたとき・フォルダでないものが同じ名前であるときは成功にする。届かない・読めないなど有無が分からないときは、試し直さずに OS の文面を添えて例外にする（権限が無いのに「削除しました」と記録しない）。フォルダとして在るのに開けないときだけ、決めた回数まで待って試し直し、使い切れば例外にする。空のパスは何もしない | [ファイルの読み書きの範囲](../../safety/file-access.md) | インデックス作成・`RemoveIndex`・`removeDroppedFolders`（インデックス・作業フォルダの削除） |
| `removeTreeNoFollow` | path | – | フォルダを中身ごと削除する。ジャンクション・シンボリックリンクは辿らず、リンクだけを外す（先は消さない）。path 自身がリンクならリンクだけを外す。先に全体を数え、リンクかどうか確かめられないフォルダ（`Unknown`）が 1 つでもあれば、何も消さずに例外を投げる。消す前に、フォルダごとにリンクでないかを確かめ直し、数えたあとにリンクに替わっていれば例外を投げる（それまでに消した分は戻らない。何も消さないと言えるのは、数える段で `Unknown` を見つけたときだけ）。確かめ直すのは、そのフォルダ自身だけで、上の階層が途中で替わったときは見つけられない。末尾の `\`・`/` は落とす。読み取り専用の属性は外してから消し、260 文字を超えるパスも扱う。フォルダが無ければ何もしない | 同上 | `removeDirectoryRetry` |
| `removeEntryOnly` | item（FileSystemInfo） | – | 1 つのファイル・フォルダ・リンクを、中を辿らずに消す（読み取り専用の属性は外してから）。フォルダは空のものだけ消す | 同上 | `removeTreeNoFollow`・`clearTmpDir` |
| `getFilesNoFollow` | root, -IncludeHidden, -SkipUnknown, -TopOnly | `@{ Files; Directories; Links; HasError; Reason }` | root の下のファイル・フォルダを、リンクのフォルダへ降りずに集める（`Get-ChildItem -Recurse`・`AllDirectories` の代わり）。ジャンクション・シンボリックリンクは `Links` に入れ、中へ降りない（リンクの先のファイルは `Files` に入らない）。root そのものはリンクかどうか見ず、先の中を集める。Hidden は、既定では除く（`-IncludeHidden` で含める。System だけのものは入れる）。`-TopOnly` を付けると root の直下だけを読み、サブフォルダの下へは進まない（直下のリンクは `Links` に入る）。並びは `Get-ChildItem -Recurse` と同じ（フォルダごとに、その中身を並べてから、サブフォルダの下へ進む）。`getLinkKind` が `Unknown` のフォルダ（OneDrive のオンデマンドのフォルダなど）は、既定では降り、`-SkipUnknown` を付けると降りずに `Links` に入れる。読めないフォルダがあっても例外にせず `HasError` を `$true`・`Reason` にして、読めたところを返す | [ファイルの読み書きの範囲](../../safety/file-access.md) | `findTargetFiles`・`getIndexTsvFilesNoFollow` |
| `getLinkKind` | item（Get-Item・Get-ChildItem などが返すもの） | string | `Link`（ジャンクション・シンボリックリンク）・`Unknown`（フォルダで `ReparsePoint` が付き種類が確かめられない、属性が読めない）・`Plain`（それ以外）のどれか | 同上 | `testLinkItem`・`removeTreeNoFollow` |
| `testLinkItem` | item（Get-Item・Get-ChildItem などが返すもの） | bool | ジャンクション・シンボリックリンクなら `$true`。属性の `ReparsePoint` が付き、`LinkType` が `Junction`・`SymbolicLink` のもの（`getLinkKind` が `Link`）。フォルダで `ReparsePoint` が付いて種類が確かめられないもの（`Unknown`）、属性が読めないときも `$true`（辿らない側に倒す）。`ReparsePoint` のファイル（OneDrive のオンデマンドなど）は `$false` | 同上 | `testPathIsLink`・`removeTreeNoFollow` |
| `testPathIsLink` | path | bool | path そのもの（途中の階層は見ない）がリンクなら `$true`。無い・空のパスは `$false`。確かめられないときは `$true`。末尾の `\` は落として調べる | 同上 | `testPathHasLink`・`clearTmpDir`・`newWorkerTmpDir`・`removeEmptyDir`・`newDirectoryReplacingExisting` |
| `testPathHasLink` | base, path | bool | base の下から path までの各階層（base 自身は見ない）に、リンクがあれば `$true`。まだ無い階層から先は見ない。path が base の下になければ、確かめられないので `$true` | 同上 | `initTmpDir`・`removeStaleTmpDirs` |
| `newDirectoryUnder` | path, root | bool | root の下に path のフォルダを（途中の階層ごと）作り、作れたら（すでにあるときも）`$true`。root から path までの途中（path を含む）にリンクがある・path が root の下でない・空のパスのときは、何も作らずに `$false`。作ったあとにもう一度確かめ、リンクに替わっていれば `$false`。root 自身は見ない（path が root と同じなら、root を作って `$true`）。作れない理由（権限など）は例外 | 同上 | `copySourceForOpen` |
| `testRemovablePath` | path, root | bool | path を root の下で消す・移す・作ってよいか。path と root が絶対パス（ドライブ・UNC）、path が root の下（`testPathUnder`）、root から path までの途中にリンクが無い（`testPathHasLink`。まだ無い階層は見ない）のすべてが成り立つときだけ `$true`。root そのものがリンクの構成は今までどおり動く。`-leafLinkOk` を付けると path そのものがリンクでもよい（途中だけ見る。`removeDirectoryUnder`・`moveDirectoryUnder` が付け、消す・移すフォルダそのものがリンクなら、リンクだけを外す・動かす）。確かめられないときは `$false` | 同上 | `removeDirectoryUnder`・`moveDirectoryUnder`・`publishIndexFiles`・`clearOldOpenCopies` |
| `removeFileUnder` | path, root | bool | `testRemovablePath` が真のときだけ、ファイル 1 つを `removeEntryOnly` で消して `$true`（path 自身がリンクならリンクだけを外す。無いときも `$true`）。偽なら何も消さずに `$false`（ログ・例外は呼ぶ側） | 同上 | `removeOldSystemIndex` |
| `removeDirectoryUnder` | path, root, tries（既定 3） | bool | `testRemovablePath` が真のときだけ `removeDirectoryRetry` で消して `$true`。偽なら何も消さずに `$false`（ログ・例外は呼ぶ側が決める）。フォルダを消す呼び出しは、これを通す（通せないものは `tests/meta/remove_entry.Tests.ps1` の一覧に理由つきで置く） | 同上 | `removeBookDir`・`RemoveIndex`・`removeOldIndexOf`・`removeOldSystemIndex`・`removeDroppedFolders`・`removeTmpDir`・`removeStaleProcessDirs`・インポート・本文インデックスの TSV の削除 |
| `moveDirectoryUnder` | source, destination, root, sourceRoot（省略時は root） | bool | 移すと元の場所から無くなるため、元（sourceRoot の下）と先（root の下）の両方を `testRemovablePath` で確かめてから `moveDirectoryRetry` で移して `$true`。偽なら何も移さずに `$false` | 同上 | `RenameIndex`・`renameOldIndexFolder`・インポート（`swapInImportedIndexDir`・`restoreSwappedIndexDir`。元と先で根が違う） |
| `moveDirectoryRetry` | source, destination, tries（既定 5）, waitMilliseconds（既定 200。試すたびに倍にする） | – | フォルダを移動する。`removeDirectoryRetry` と同じ理由で、待つ時間を倍にしながら数回試す | [インポート](../index-data/format.md#インポートimportindex) | インデックスのインポート（`swapInImportedIndexDir`・`restoreSwappedIndexDir`）・ワークスペースの移動（`WorkspaceMover.MoveEntries`） |
| `testSameDriveRoot` | first, second | bool | 2 つのパスが同じドライブ（同じ根）にあるか。同じなら `Directory.Move`、違えば写してから消す。テストが別のドライブを用意せずに差し替えられるよう関数にしてある | – | `WorkspaceMover.MoveEntry` |
| `testAnyEntry` | entries（`EnumerateFiles` などの列挙）, predicate（省略可） | 真偽値 | 条件に合うものが 1 つでもあるか。見つけたところで列挙をやめ、列挙子を必ず Dispose する（`foreach` を途中で抜けると下のフォルダを掴んだまま残り、そのフォルダを移動・削除できなくなるため） | – | 前の版のワークスペースの状態（`getLegacyIndexState`）・インデックスの有無（`testIndexExists`）・元のファイルごとのフォルダの判定（`testIndexBookDir`） |
| `selectFirstEntries` | entries（`EnumerateFileSystemEntries` などの列挙）, count | 配列 | 列挙の先頭から count 件までを配列にして返す。count 件で列挙をやめ、列挙子を必ず Dispose する（`testAnyEntry` と同じ理由） | – | 作業フォルダに選べるかの判定（`testDefaultWorkspace`・`getFolderEntrySample`） |

## TSV の整形・ファイル名の確かめ（`shared/core/text.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `replaceCellNewLine` | inputString | string | `"` で囲まれた範囲の改行（CRLF・CR・LF）を `$cellNewLine` に置き換える | [Excel](../indexing/excel.md) | formatTsv |
| `formatTsv` | content, firstRow（既定 1）, firstColumn（既定 1） | string | TSV 整形。N 行目・k 列目をシートの N 行目・k 列目にそろえる | 同上 | prettyTsv |
| `countTsvFields` | line | int | Excel に貼り付けたときのセル数（`"` で始まるセルは閉じる `"` までを 1 セルとする。先頭のセルが空でも数え落とさない） | [1 行の組み立て](../search/output.md#1-行の組み立て) | 検索 |
| `toColumnName` | number | string | 列番号を列名に変換（1 → `A`、27 → `AA`） | 同上 | toResultHeader |
| `splitTsvCells` | line | string[] | TSV の 1 行をセルに分ける（`"` で囲まれたセルは 1 セルとし、囲みを外す。`countTsvFields` と同じ区切り方。画面のプレビューは同じ区切り方を型 `HitRow`（`ui/search/result_types.ps1`）の中に持つ） | – | テストだけ |
| `testTextTrimmed` | actualWidth, requiredWidth, tolerance | bool | 文字が枠に入りきらず省略（…）で切れているか（要る幅が実際の幅より誤差を超えて大きいときだけ $true。幅 0 以下は $false） | 同上 | addTrimmedToolTip（`shared/ui/trimmed_tooltip.ps1`） |
| `toSafeFileName` | name | string | ファイル名禁止文字を全角に置換 | [インデックス作成](../indexing/index.md) | インデックス名（newIndexName・getTargetFolders）、シート名の照合（画面。元のファイルを開くとき） |
| `getFolderNameProblemText` | name, allowReserved（既定 `$false`） | string（問題の文。使えれば空） | フォルダ名として使えるか（空・前後の空白・長すぎる・使えない文字・最後の `.`・予約名）。allowReserved が `$true` のときは予約名（`aux.backup` など）を拒まない（既にある実在のフォルダの名前を空にしないため。新しく付ける名前には使わない）。どのツールからも使う | 同上 | testIndexName, testStoredIndexName, newIndexName, getTargetFolders, readStatusFile |
| `testPlainFolderName` | name | bool | 外へ出ず名前のとおりに作れるフォルダ名か（空・`.`・`..`・長すぎる・使えない文字だけを拒む。予約名・最後の `.` や空白は問わない） | 同上 | readStatusFile（`LegacyFolders` の判定） |
| `formatDisplayText` | text, keepLineBreaks（スイッチ） | string | インデックスから来た文字列を画面に出すときの表示用の文字列。制御文字（C0・C1）と向きを変える文字（U+202E など。並びは `$displayControlChars`）を U+FFFD にする。`keepLineBreaks` ではタブ・改行を残す。コピー・開くには元の値を使う | [元のファイルを開く](../gui/open-file.md) | newFileGroup, addFileGroupLocation, openFoundSource ほか結果の表の表示 |

日時の文字列は `formatTime`（time → string。当日なら `H:mm`、それ以前は `M/d H:mm`。`$null` は空。判断層の `search_bar_view.ps1` と画面層が使う）も `text.ps1` にある。

## TSV のファイルの書き出し（`shared/core/tsv_file.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `prettyTsv` | 入力パス, 出力パス, firstRow（既定 1）, firstColumn（既定 1） | bool | 入力（UTF-16）を読み、`formatTsv` して UTF-8（BOM 付き）で保存。内容が空なら保存せず `$false` | 同上 | インデックス作成 |

## テキストファイルの読み取り（`shared/core/text_file.ps1`）

詳細は [テキストファイルの読み取り](../indexing/text.md)。バイト列から判定する関数はファイルを読み書きしない（`readTextFile` だけがファイルに触る）。

| 関数 | 入力 | 出力 | 概要 | 使用元 |
|---|---|---|---|---|
| `testTextExtension` | path | bool | 拡張子が対象のテキストの拡張子（`$textExtensions`。大文字・小文字を区別しない）か | ingestFile, getContentIndexFileKind, describePlace, findTargetFiles（`$targetExtensions`） |
| `testTextOpenWithNotepad` | path | bool | 拡張子が、既定のアプリではなくメモ帳で開く拡張子（`$textNotepadExtensions`。開くと実行・登録になるもの。大文字・小文字を区別しない。終わりの `.` と空白は Windows が無視して開くので、落としてから見る）か | openFoundSource（[元のファイルを開く](../gui/open-file.md)） |
| `detectTextEncoding` | bytes | string / `$null` | バイト列だけから文字コード（`UTF8` / `UTF16LE` / `UTF16BE` / `ShiftJIS` / `EUCJP` / `ISO2022JP`）を判定する。判定できない・あいまいなものは `$null`（取り込まない） | readTextFile |
| `detectJapaneseUtf16WithoutNul` | bytes | string / `$null` | NUL の無い BOM 無し UTF-16（日本語だけの文章）を、かなの割合と日本語の文章に出る文字の割合で判定する（`UTF16LE` / `UTF16BE`） | detectTextEncoding |
| `getIso2022JpVerdict` | bytes | string | ISO-2022-JP かの判定（`yes` / `broken`（ESC $ B があるが 8 ビット・規格外の ESC がある）/ `none`） | detectTextEncoding |
| `detectLegacyJapaneseEncoding` | bytes | string / `$null` | UTF-8 として読めなかったバイト列が Shift_JIS か EUC-JP か。決まらなければ `$null` | detectTextEncoding |
| `testHalfWidthKanaNatural` | text | bool | 半角カナの多い文字列が、濁点・半濁点の位置など半角カナの並びとして自然か | detectLegacyJapaneseEncoding |
| `tryDecodeStrict` | bytes, codePage | string / `$null` | 読めないバイト列があれば `$null`、読めれば文字列 | detectTextEncoding |
| `testTextPlausible` | text | bool | 制御文字・私用領域・U+FFFD を含まない（テキストとして自然）か | detectLegacyJapaneseEncoding |
| `decodeTextBytes` | bytes, encodingName | string | `detectTextEncoding` が返した文字コードで、バイト列を文字列にする（BOM は取り除く） | readTextFile |
| `splitTextLines` | text | string[] | `StreamReader.ReadLine` と同じ分け方（CRLF・LF・CR）で行に分ける。途中の空の行は残し、行末の空白は取り除き、末尾の空の行は捨てる | readTextFile |
| `readTextFile` | path, maxBytes（既定 `$textFileMaxBytes`） | string[] | ファイルを読み取り専用の共有で開き、大きさの上限・文字コードを確かめてから行の並びにする。上限超え・バイナリは例外（[エラーメッセージ一覧](../indexing/errors-ingest.md#ファイルごとの失敗取り込み一覧のエラー列)） | extractTextFile |

## インデックスへの配置（`tebunko/index/index_store.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `getIndexTsvCounts` | dir（既定 `$workspace.IndexDir`） | 相対パス → 数の辞書 / `$null` | インデックスの中のファイルを 1 回列挙して数える。本文インデックスのファイルはそのファイルの相対パス → 1（0 バイトなら -1）、本文インデックスに入れる前の TSV は元のファイルのフォルダの相対パス（取り込み一覧の相対パスと同じ）→ TSV の数（0 バイトの TSV があれば -1。TSV の無いフォルダは 0）。列挙できなければ `$null` | [取り込み対象の決定](../indexing/target-decision.md#取り込み対象の決定createtargetlist) | インデックス作成（取り込み対象の決定） |
| `testIndexComplete` | row, relPath, counts | bool | 取り込み一覧の「済」の行に対して、インデックスがそろっているかを返す。そのフォルダ・その拡張子の本文インデックスのファイルがあれば（0 バイトでなければ）`$true`。無ければ本文インデックスに入れる前の TSV を見て、行の TSV 数より少ない・0 バイトなら `$false`（取り込み直す）。TSV 数が空・counts が `$null` のときは確認しない | 同上 | インデックス作成（取り込み対象の決定） |
| `publishIndexFiles` | fromDir, bookDir, stagingDir, indexRoot, publishRoot | – | 書き出した TSV を stagingDir に集めてから、`bookDir` をフォルダごと入れ替える（作りかけのインデックスを残さない）。別ドライブでフォルダごと移せない場合は 1 件ずつ移す | [入れ替えと書き出し](../index-data/publish.md#インデックスへの入れ替えpublishtsv--publishindexfiles) | インデックス作成（publishTsv） |

## 本文インデックスの形式（`tebunko/index/content_index_format.ps1`）

形式は [インデックスのファイルの形](../index-data/format.md#配置命名規則)「本文インデックスの形式」。判断層のため、ファイルを読み書きしない。

| 関数 | 入力 | 出力 | 概要 | 使用元 |
|---|---|---|---|---|
| `getContentIndexFileKind` | book | string | 元のファイル名から種類（`Excel` / `Word` / `PowerPoint` / `テキスト` / `PDF`。分からなければ空） | convertBookToContentIndexBlock, convertPlaceToContentIndexMeta, describePlace, toResultLine, addFileGroupLocation・ensureRows（getPlace に渡す値・ファイルの見出しの `AppKind`・検索結果の行の `IsExcel`）, testCellPrefixedPlace, searchContentIndexFiles（長い行を切るかの判定）, readContentIndexLines |
| `getContentIndexExtension` / `getContentIndexFileName` / `readContentIndexFileName` | book / extension, part / name | string / `@{Extension; Part}` | 本文インデックスを分ける拡張子（小文字・`.` なし） / 本文インデックスのファイルの名前（`content_index.<拡張子>.<番号>.tsv`） / 名前から拡張子と番号を取り出す | convertFolderToContentIndex, getIndexTsvCounts |
| `splitContentIndexBooksByExtension` | books | [ordered] 拡張子 → 並び | 元のファイルの並びを拡張子ごとに分ける（各並びの中の順は変えない） | convertFolderToContentIndex |
| `encodeContentIndexValue` / `decodeContentIndexValue` | value | string | メタ情報の値の制御文字（タブを除く）と `%` を `%XX` にする / 戻す | convertToContentIndexText, readContentIndexPlaces |
| `convertPlaceToContentIndexMeta` | book, place | [ordered] キー → 値 | 場所の名前（TSV のファイル名。`見積[図形]`・`見積[ヘッダー・フッター]`・`ページ001` など）を場所のメタ情報にする。組み立て直して同じ名前にならないものは `部分=<名前>`・`対象=本文` | convertToContentIndexText |
| `convertContentIndexMetaToPlace` | meta | string | 場所のメタ情報から場所の名前を組み立てる（画面の表示・図形とコメントの除外・元のファイルを開く処理が使う形） | readContentIndexPlaces |
| `convertToContentIndexBody` | text | string | TSV の中身を本文インデックスに入れる形にする（改行を LF に、末尾に LF、U+001C〜U+001F を除く） | convertToContentIndexText |
| `convertToContentIndexText` | books, version（既定 `contentIndexVersion`） | string | 元のファイルの並び（TSV から新しく作るもの、または前の本文インデックスから写すまとまり）から、本文インデックスの文字列を作る。1 行目の版は version（分けた塊のファイルは `contentIndexSplitVersion`） | convertFolderToContentIndex, writeContentIndexPart |
| `getContentIndexPlaceMetaText` | book, place | string | 場所 1 つ分のメタ情報の行（`<RS> キー=値` を LF で並べたもの） | convertBookToContentIndexBlock, measureContentIndexPlaceChars, splitContentIndexBookBlock |
| `splitContentIndexBookBlock` | book（`@{Name; Places}`）, maxBytes（既定 `contentIndexSearchMaxBytes`）, chunkBytes（既定 `contentIndexFileMaxBytes`） | string の並び | 元のファイル 1 つの、本文インデックスに入れるまとまりを返す。maxBytes 以下なら分けずに 1 つ。超えるなら約 chunkBytes の塊に分ける（場所の切れ目を優先し、1 つの場所が 1 つの塊に入らないときだけ行で切って、続きの塊に `開始行` を付ける）。塊が maxBytes におさまらない（1 行が大きすぎる）ときは例外 | [分けた塊](../index-data/format.md#分けた塊版-3) | readNewContentIndexBooks |
| `addContentIndexPlaceToChunks` | chunks, current, head, meta, body, limit | – | `splitContentIndexBookBlock` の、場所 1 つ分（開いている塊に入るなら足し、入らなければ閉じて新しい塊から始め、1 つの塊に入らない場所は行で切る） | 同上 | splitContentIndexBookBlock |
| `readContentIndexPlaces` | text | `@{Book; Location; CellPrefixed; FirstLine; Start; End}` の並び | 本文インデックスの文字列から、場所ごとの元のファイル名・場所の名前・行の先頭がセル番地か（`CellPrefixed`。Excel の図形・コメントが `$true`。`testContentIndexCellPrefixed`）・場所の最初の行が場所の何行目か（`FirstLine`。分けた塊の `開始行`。ふつうは 1）・中身の範囲を先頭から順に返す。版が 2・3 でなければ例外 |
| `addContentIndexPlace` | places, current, meta, end | – | `readContentIndexPlaces` の、場所 1 つを閉じる処理（終わりの位置・場所の名前・セル番地の判定・開始行を決めて places に足す。開いてよい名前でない場所は足さない） | 同上 | readContentIndexPlaces（検索のスレッドにも読み込む） | 検索（searchContentIndexFiles）、readContentIndexLines |
| `testContentIndexCellPrefixed` | book, meta | bool | 場所の行が「セル番地 + タブ + 文字」の形か。メタ情報の `対象` を `testCellPrefixedPlace` に渡して返す | readContentIndexPlaces |
| `testCellPrefixedPlace` | book, kind | bool | 場所の行が「セル番地 + タブ + 文字」の形か（元のファイルが Excel で、場所の種類が図形・コメント。ふつうの場所の種類は空）。この判定はここだけにあり、検索結果の行（`HitRow.IsCellPrefixed`）も行を作るときにこれで決める | testContentIndexCellPrefixed、検索結果の行を作る所（result_list.ps1） |
| `splitContentIndexTextByBook` | text | `@{Name; Block}` の並び | 本文インデックスの文字列を元のファイルごとのまとまりに分ける（入れ替えない元のファイルをそのまま写すため）。版が違えば例外 | convertFolderToContentIndex |
| `getContentIndexBodyText` | text | string | メタ情報の行を除いた中身（2-gram の値を作るため） | getFileGramValues |

## インデックスの格納形式の目録（`tebunko/index/content_index_format.ps1`・`index_format_store.ps1`）

形式は [インデックスのファイルの形](../index-data/format.md#置き場と目録store)「置き場と目録」。判定は判断層（`getIndexFormatState`）、ファイルの読み書きは状態層（`index_format_store.ps1`）。

| 関数 | 入力 | 出力 | 概要 | 使用元 |
|---|---|---|---|---|
| `getIndexFormatState` | catalog, oldNames, v2Names, hasLegacyIndex | `@{State; Message; Indexes}` | 目録の行と、古い置き場・今の置き場にあるインデックスの名前から、ワークスペースと各インデックスの状態（Current・Building・Rebuilding・Old・Empty・Broken・Newer）・検索に出すか・文言を決める。ファイルには触れない | getWorkspaceIndexFormat |
| `testContentIndexVersionLine` | line | bool | 本文インデックスの先頭行が読める版（2、分けた塊のファイルの 3）の `版=` の行か | testIndexFileCurrentVersion |
| `getContentIndexTextVersion` | text | int | 本文インデックスの文字列の版（1 行目が `版=3` なら `contentIndexSplitVersion`、それ以外は `contentIndexVersion`） | 同上 | readPreviousContentIndexParts, readContentIndexHeadPlace |
| `readIndexFormat` / `writeIndexFormat` | path / indexes, path, appVersion | `@{Status; Version; AppVersion; Indexes}` / なし | 目録を読む（Ok・Missing・Broken・Newer・Older）/ 一時ファイルに書いて入れ替える | getWorkspaceIndexFormat, updateIndexFormat |
| `updateIndexFormat` | change, path | bool | 目録を排他で読み、change で書き換えて書く。今より新しい版は書き換えない | setIndexFormatRow, removeIndexFormatRow, renameIndexFormatRow, markIndexFormatBuilding, completeIndexFormat |
| `setIndexFormatRow` / `removeIndexFormatRow` / `renameIndexFormatRow` | name, state / name / oldName, newName | bool | インデックス 1 つの行を書く / 消す / 名前を変える | index_store.ps1, index_catalog.ps1 |
| `repairIndexFormat` | workspace | 目録の読み取り結果 | 目録が壊れたとき、今の版の本文インデックスのファイルがあるインデックスを「作成中」として書き直す | getWorkspaceIndexFormat |
| `getWorkspaceIndexFormat` | workspace | `getIndexFormatState` の結果 | 今のファイルから、ワークスペースと各インデックスの状態を求める（目録が壊れていれば先に直す） | indexer_run.ps1, index_detail.ps1, index_list.ps1 |
| `findOldIndexNames` / `getIndexNamesWithContent` | workspace / dir | 名前の並び | 古い置き場・今の置き場で本文インデックスを持つインデックスの名前 | getWorkspaceIndexFormat |
| `markIndexFormatBuilding` | names, workspace | bool | この回で作るインデックスのうち目録に行が無いものを「作成中」で足す | invokeIndexerBody |
| `completeIndexFormat` | names, workspace | `@{Updated; Failed}` | 作り終えたインデックスの行を「作り終えた」にし、古い置き場の分を消す | invokeIndexerBody |
| `removeOldIndexOf` | name, workspace | `@{Removed; Failed}` | 古い置き場の `<名前>` と、空になった古い置き場（`content_index\`・古い `ingest_status.tsv`）を消す。名前が不正なら何も消さない。ジャンクションは辿らない | completeIndexFormat |

## 本文インデックスの読み書き（`tebunko/index/content_index_store.ps1`）

| 関数 | 入力 | 出力 | 概要 | 詳細 | 使用元 |
|---|---|---|---|---|---|
| `writeContentIndexFile` / `readContentIndexText` | path, text / path | – / string | 本文インデックスのファイルを UTF-16LE（BOM 付き）で書く（`<名前>.tmp` に書いてから `File.Replace` で置き換える） / 読む（置き換え・削除を妨げない共有モード） | [インデックスのファイルの形](../index-data/format.md#配置命名規則) | convertFolderToContentIndex, readContentIndexLines |
| `getIndexFolderBooks` | folder | `@{Name; Places}` の並び | フォルダ直下の元のファイルごとのフォルダ（`<ファイル名.xlsx>\<場所>.tsv`）から、本文インデックスに入れる元のファイルと場所の並びを作る（ファイル名は `convertIndexFileNameToPlace` で場所に戻す。並びは前の版と同じになるよう、戻した場所を符号化し直したキーで並べ替える） | 同上 | convertFolderToContentIndex |
| `measureContentIndexPlaceChars` | book, place, text | long | 場所 1 つ分（メタ情報の行＋中身）の本文インデックスの文字数（`convertBookToContentIndexBlock` と同じ作り方） | 同上 | assertContentIndexBookFits |
| `getContentIndexTooLargeMessage` / `getContentIndexDroppedText` | – / rel, count | string | 1 つの元のファイルで検索の上限を超えるときの取り込み一覧の文言 / 上限超えのファイルを外したときの黄色のログの 1 行 | [取り込み一覧のエラー](../indexing/errors-ingest.md) | assertContentIndexBookFits, publishIndexFolders |
| `assertContentIndexBookFits` | dir, bookName, maxBytes | – | 作業フォルダの TSV（元のファイル 1 つぶん）の 1 行が、それだけで本文インデックスのファイルを検索の上限より大きくするなら例外にする（元のファイルが大きいだけなら塊に分けて入れるので通す。バイト数から届かないと分かれば読まない） | [本文インデックスへの書き出し](../index-data/publish.md) | publishTsv |
| `readNewContentIndexBooks` / `readPreviousContentIndexParts` / `getOversizeRequeueNames` | books / longDest, keep / path, keep | `@{Books; TooLarge; Tsv}` / `@{Parts; Texts; Requeue; Remove; Dropped}` / string[] | 入れる元のファイルの TSV を読み、1 つで上限を超えるものは塊に分ける（`Books` の `Chunks`）。分けても入らないものは `TooLarge` に入れる / 前の本文インデックスのファイルを読み（`Parts` の各要素に `Version`）、上限を超えるファイルは名前だけ流れで読んで消す対象と取り込み直しにする | 同上 | convertFolderToContentIndex |
| `findOversizeContentIndexFolders` | root | string[] | 本文インデックスのファイルが検索の上限を超えるフォルダ（フルパス）を返す（ファイルの大きさだけで見る） | 同上 | recoverIngestJournal |
| `convertFolderToContentIndex` | folder, destFolder, removeBooks, removeTsv | `@{Books; Tsv; Chars; Files; Texts; Requeued; Dropped}` | フォルダ 1 つの TSV から拡張子ごとの本文インデックスのファイルを書く。前の本文インデックスのファイルとまぜ（TSV のある元のファイルは入れ替え、removeBooks は外し、ほかは写す）、元のファイルが無くなった拡張子の本文インデックスのファイルは消す。removeTsv なら書き終えた後に TSV のフォルダを消す。Texts は書いた中身 | 同上 | updateFolderContentIndex |
| `writeContentIndexPart` | part, destFolder, oldTexts | string | 計画（`planContentIndexParts`）の本文インデックスのファイル 1 つを書き（書き直しが要らなければ書かずに前の中身を返す）、中身を返す。1 行目の版は part の `Version` | 同上 | convertFolderToContentIndex |
| `updateFolderContentIndex` | folder, removeBooks | 同上 | `convertFolderToContentIndex` を同じフォルダに書き、TSV を消す形で呼ぶ | 同上 | publishIndexFolders |
| `findContentIndexFiles` | root, relPath, recurse | `@{Path; Root; RelDir; RelPath; Ticks; Size}` の配列 | フォルダ以下の本文インデックスのファイルを列挙し、フォルダの順・フォルダの中は名前の順に並べる（Path は `\\?\` 付き。Ticks・Size は読んだ内容を使い回してよいかの判定に使う） | [検索を速くする仕組み](../search/speed.md) | getContentIndexFiles |
| `findIndexFoldersWithBooks` | root | string[] | 元のファイルごとのフォルダ（本文インデックスに入れる前の TSV）が直下にあるフォルダを返す（インデックス作成が途中で止まったとき） | [インデックスのファイルの形](../index-data/format.md#配置命名規則) | インデックス作成（開始時） |
| `publishIndexFolders` | pending（フォルダ → 無くなった元のファイル名）, indexRoot | `@{ Count（書き出したフォルダの数）; Skipped（飛ばしたフォルダの `@{ Rel; Stage; Reason }`）; Published（本文インデックスのファイルを書き終えたフォルダ）; Requeued（取り込み直しにする元のファイルの、インデックスからの相対パス）}` | フォルダごとに、本文インデックスに書き（`updateFolderContentIndex`）、TSV を消し。フォルダの 1 つが失敗しても例外にせず、そのフォルダを飛ばして残りを続け、ログに黄色で出す（本文インデックスを書き終える前なら何も書き換えない） | 同上 | インデックス作成（indexer_run.ps1 の flushPending） |
| `readContentIndexLines` | path, book, location, firstLine, count, cache, marks | `@{Rows; End; Marks}` | 本文インデックスのファイルの中の元のファイル book・場所 location の firstLine 行目から count 行を返す（行の数え方は検索と同じ。`Rows` は `@{LineNumber; Line}` の並び。`End` は count 行に満たなかった＝その先にもう行が無い）。cache（`newTsvTextCache`）に同じ本文インデックスのファイルの内容があれば読み直さない。無ければ、`marks`（前に読んだときの控え。ファイルのパス・更新日時・サイズ・場所の一覧・場所ごとの 200 行ごとの文字の位置）が今のファイルのものである間は、場所の一覧を使い回し、控えのうち目当てに一番近い所から必要な行までだけを読む（`readContentIndexRange`）。`marks` が無い・別のファイルのもの・更新日時かサイズが違うとき、範囲を読めないとき、読んでいる間にファイルが置き換わったときは、ファイルの全体を読む（初めて開くファイルの最初の読みはこれに当たる）。更新した控えを `Marks` で返す（引数で渡し戻り値で受ける。複数のスレッドで動くため `$script:` に持たない。渡された控えは書き換えない）。更新日時もサイズも同じまま中身だけが書き換わったファイルは、控えが古いことを見分けられない。読めない・見つからなければ空 | [選択行のプレビュー](../gui/preview.md) | readContentIndexLinesAcrossParts |
| `readContentIndexLinesAcrossParts` | path, book, location, firstLine, count, cache, marks | `@{Rows; End; Marks}` | `readContentIndexLines` と同じだが、場所が複数の塊のファイル（版 3）に分かれているときは、前の塊・続きの塊のファイルも読み、行の番号を場所の中の通しの番号で返す。分けたファイルでなければ `readContentIndexLines` をそのまま呼ぶ | [選択行のプレビュー](../gui/preview.md) | 画面 |
| `readContentIndexHeadPlace` | path, chars（既定 8192） | `@{Split; Place}` / null | 本文インデックスのファイルの先頭だけを読み、分けた塊のファイルか（版 3）と、最初の場所（`readContentIndexPlaces` の 1 件）を返す。読めなければ `$null` | 同上 | readContentIndexLinesAcrossParts, getContentIndexSplitChain |
| `getContentIndexSplitChain` | path, book, location, head | `@{Path; FirstLine}` の並び | 分けた場所のすべての塊のファイル（同じフォルダ・同じ拡張子で、先頭の場所が同じ元のファイル・場所のもの）を、場所の中の位置の順に返す | 同上 | readContentIndexLinesAcrossParts |
| `getContentIndexBookHeadPrefix` | – | string | 本文インデックスで元のファイル 1 つの始まりを表す行の、値の前までの部分（`<RS> ファイル名=`）。書き出し・分割・名前の取り出しが同じ書式を使うための 1 か所 | 同上 | `convertToContentIndexText`・`splitContentIndexTextByBook`・`readContentIndexBookNamesStreaming` |
| `readContentIndexBookNamesStreaming` | path, maxLineChars（既定 32768） | string の配列 | 本文インデックスのファイルに入っている元のファイルの名前（`ファイル名=` の行の値）を、全体を読み込まずに区切り読みで集める。上限を超える行は改行まで読み捨てる。読めなければ空の配列。エクスポートが、大きすぎて含めなかったファイルの元のファイルを知らせるために使う | `writeContentIndexFilesToArchive` |
| `readContentIndexRange` | path, start, end, firstLine, lines, baseLine（既定 1）, checkpoints | string / null | 本文インデックスのファイル（UTF-16LE・BOM 付き）の、文字の位置 `start`（含む）から `end`（含まない）までを `start` から数え、`firstLine` 行目の先頭から `lines` 個目の改行までを文字列で返す（それより前の行は持たない）。`baseLine` は `start` の行の番号。`checkpoints`（ハッシュテーブル）を渡すと、通った 200 行ごとの行の先頭（行の番号 → 文字の位置）を足す。BOM が無ければ `$null`（呼び出し側がファイルの全体を読む）。置き換え・削除を妨げない共有モードで開く | 同上 | readContentIndexLines |

