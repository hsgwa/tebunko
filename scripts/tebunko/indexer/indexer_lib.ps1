# インデックス作成の部品の読み込み口。indexer.ps1 と、取り込みのスレッド（indexer_run.ps1 の ingestWorkerScript）が dot-source する。
# 画面は読み込まない（インデックス作成は、画面のインデクサのスレッドが indexer.ps1 を実行して行う）
. ([WildcardPattern]::Escape("$PSScriptRoot\..\lib.ps1"))
. ([WildcardPattern]::Escape("$PSScriptRoot\..\..\shared\office\office_reader.ps1"))
. ([WildcardPattern]::Escape("$PSScriptRoot\..\..\shared\office\office_embedded.ps1"))
. ([WildcardPattern]::Escape("$PSScriptRoot\..\..\shared\office\office_app.ps1"))
. ([WildcardPattern]::Escape("$PSScriptRoot\pending_publish.ps1"))
. ([WildcardPattern]::Escape("$PSScriptRoot\indexing_reporter.ps1"))
. ([WildcardPattern]::Escape("$PSScriptRoot\indexer_plan.ps1"))
. ([WildcardPattern]::Escape("$PSScriptRoot\extract_office.ps1"))
. ([WildcardPattern]::Escape("$PSScriptRoot\extract_text.ps1"))
. ([WildcardPattern]::Escape("$PSScriptRoot\index_migrate.ps1"))
. ([WildcardPattern]::Escape("$PSScriptRoot\indexer_run.ps1"))
