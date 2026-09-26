# Markdown のリンク切れのテスト。
# git で管理している .md すべての相対リンクとアンカーを tools\check_markdown_links.ps1 で調べる。
# コミット前のフック（.md を変えたとき）と CI の必須チェック test で止め、リンクが切れたままマージできないようにする。
# docs\ の中は mkdocs build --strict（docs.yml）でも調べるが、docs\ の外と docs\ から外へのリンクはここでしか調べない。
BeforeAll {
    . "$PSScriptRoot\..\helpers\load.ps1"
}

Describe "Markdown のリンク" -Tag Meta {
    It "リンク先のファイルと見出しがある" {
        $output = @(& "$here\..\tools\check_markdown_links.ps1" 6>&1 | ForEach-Object { "$_" })
        # 切れたリンクの一覧を失敗の理由に出す
        (@($output | Where-Object { $_.StartsWith("  ") }) -join "`n") | Should -Be ""
        $LASTEXITCODE | Should -Be 0
    }
}
