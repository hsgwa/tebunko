# tebunko_diff（比較）の画面以外の部品の読み込み口。
# 画面・抽出プロセス（differ.ps1）・テスト・画面が起こす別スレッドから dot-source して使う。読み込む順に意味がある。
. "$PSScriptRoot\..\shared\shared.ps1"
. "$PSScriptRoot\core\paths_diff.ps1"
. "$PSScriptRoot\core\settings_diff.ps1"
. "$PSScriptRoot\diff\diff_core.ps1"
. "$PSScriptRoot\diff\diff_office.ps1"
. "$PSScriptRoot\diff\diff_folder.ps1"
. "$PSScriptRoot\diff\diff_report.ps1"
. "$PSScriptRoot\job\diff_job.ps1"
