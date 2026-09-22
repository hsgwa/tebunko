# 制限モードの部品の読み込み口（start.ps1 が、制限モードで起動するときに読み込む）。
#
# 制限モードも、いつもの画面と同じ判断（設定・インデックス名・検索の条件・元のファイルの場所）を使う。
# ここから読み込むファイルには、いつもの画面だけが使う関数（Excel の COM・画面の部品など）も入っている。
# 制限モードが使う関数は、制限言語モードでも動く書き方にしてある。tests/meta/clm.Tests.ps1 が、
# 同じ呼び出しを FullLanguage と模擬の制限言語モードで動かし、結果が同じことを確かめる（tests/meta/clm_cases.ps1）。
# 読み込む順に意味がある（paths.ps1 が ${fullLanguage}・${rootDir} を決める）。
. "$PSScriptRoot\..\..\shared\core\paths.ps1"
. "$PSScriptRoot\..\..\shared\core\fs.ps1"
. "$PSScriptRoot\..\..\shared\core\text.ps1"
. "$PSScriptRoot\..\..\shared\core\folder.ps1"
. "$PSScriptRoot\..\core\paths_grep.ps1"
. "$PSScriptRoot\..\core\settings_grep.ps1"
. "$PSScriptRoot\..\index\index_name.ps1"
. "$PSScriptRoot\..\index\index_store.ps1"
. "$PSScriptRoot\..\search\search_query.ps1"
. "$PSScriptRoot\..\search\source_map.ps1"
