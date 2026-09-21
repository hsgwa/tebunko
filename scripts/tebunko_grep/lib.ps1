# tebunko_grep の画面以外の部品の読み込み口。
# 画面（gui）・変換処理・テスト・画面が起こす別スレッドから dot-source して使う。読み込む順に意味がある。
. "$PSScriptRoot\..\shared\shared.ps1"
. "$PSScriptRoot\core\paths_grep.ps1"
. "$PSScriptRoot\core\settings_grep.ps1"
. "$PSScriptRoot\index\index_name.ps1"
. "$PSScriptRoot\index\index_store.ps1"
. "$PSScriptRoot\indexer\indexer_state.ps1"
. "$PSScriptRoot\indexer\indexer_decide.ps1"
. "$PSScriptRoot\search\search_query.ps1"
. "$PSScriptRoot\search\search_run.ps1"
. "$PSScriptRoot\search\source_map.ps1"