# 共通基盤（どのツールからも使う部品）の読み込み口。
# ツール側は、このファイルを dot-source すれば共通基盤がそろう。読み込む順に意味がある。
. "$PSScriptRoot\core\paths.ps1"
. "$PSScriptRoot\core\fs.ps1"
. "$PSScriptRoot\core\worker_pool.ps1"
. "$PSScriptRoot\core\data_dir.ps1"
. "$PSScriptRoot\core\text.ps1"
. "$PSScriptRoot\core\folder.ps1"
. "$PSScriptRoot\core\version.ps1"
. "$PSScriptRoot\office\office_files.ps1"
. "$PSScriptRoot\office\office_process.ps1"