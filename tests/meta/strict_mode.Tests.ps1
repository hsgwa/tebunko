# テストの StrictMode（Version 2）の決まりのテスト。
# 打ち間違えた変数名が $null で素通りしてテストが緑になるのを防ぐ。例外の一覧は tests\helpers\strict_mode_exceptions.ps1
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
    . "$PSScriptRoot\..\helpers\strict_mode_exceptions.ps1"
    $script:rootDir = (Resolve-Path "$PSScriptRoot\..\..").Path
}

Describe "テストの StrictMode" -Tag "Meta" {
    It "例外の一覧は、全部に理由が付いている" {
        $noReason = @(getStrictModeExceptions | Where-Object { [string]::IsNullOrWhiteSpace($_.Reason) })
        $noReason.Count | Should -Be 0
    }

    It "一覧のファイルと Set-StrictMode -Off を書いたテストファイルが一致する" {
        $listed = @(getStrictModeExceptions | ForEach-Object {
            Test-Path -LiteralPath (Join-Path $script:rootDir $_.Path) | Should -BeTrue
            ([string]$_.Path).Replace("\", "/")
        } | Sort-Object)
        $testsPrefix = (Join-Path $script:rootDir "tests").TrimEnd('\') + '\'
        $written = @(Get-ChildItem -LiteralPath (Join-Path $script:rootDir "tests") -Recurse -Filter "*.Tests.ps1" |
            Where-Object { Select-String -LiteralPath $_.FullName -Pattern '^\s*Set-StrictMode\s+-Off\b' -Quiet } |
            ForEach-Object { "tests/" + $_.FullName.Substring($testsPrefix.Length).Replace("\", "/") } | Sort-Object)
        ($written -join "`n") | Should -Be ($listed -join "`n")
    }
}
