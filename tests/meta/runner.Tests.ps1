# テストの実行口（tests\run.ps1）のテスト
. "$PSScriptRoot\..\helpers\load.ps1"

Describe "テストの実行口" -Tag Meta {
    # pre-commit フックと同じく powershell.exe -File で呼ぶ。-File ではカンマ区切りのタグが 1 つの文字列で渡る。
    # 子のプロセスでは、$TestDrive に置いた 1 件だけのテストを動かす（本物のテストを全部読み込むと遅いため）
    $runner = "$here\run.ps1"
    $sample = "$TestDrive\sample.Tests.ps1"
    Set-Content -LiteralPath $sample -Encoding UTF8 -Value @(
        'Describe "見本" -Tag Unit {'
        '    It "通る" { 1 | Should Be 1 }'
        '}'
    )
    function invokeRunner([string]$tag) {
        & powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File $runner -Path $sample -Tag $tag -Quiet | Out-Null
        $LASTEXITCODE
    }

    It "実行したテストが 0 件なら失敗にする" {
        invokeRunner "NoSuchTag" | Should Be 1
    }

    It "-File で渡したカンマ区切りのタグを分けて受け取る" {
        invokeRunner "NoSuchTag,Unit" | Should Be 0
    }
}
