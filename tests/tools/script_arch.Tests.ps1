# スクリプトの依存を AST で読む部品（tools\script_arch.ps1）のテスト。$TestDrive の小さな見本で、検出するものと検出しないものを確かめる
BeforeAll {
    . "$PSScriptRoot\..\..\tools\script_arch.ps1"

    # 見本の scripts フォルダを作る（<相対パス> = <本文> の表から）。作ったフォルダのパスを返す
    function newSampleRoot {
        param ([hashtable]$files)

        $root = Join-Path $TestDrive ([guid]::NewGuid().ToString("N"))
        foreach ($key in $files.Keys) {
            $path = Join-Path $root $key
            New-Item -ItemType Directory -Force -Path (Split-Path $path -Parent) | Out-Null
            [System.IO.File]::WriteAllText($path, $files[$key], (New-Object System.Text.UTF8Encoding($false)))
        }
        return $root
    }

    # 見本の依存を "<from> -> <to>（<名前>）, ..." の 1 つの文字列で返す
    function dependenciesOf {
        param ([hashtable]$files, [string[]]$from, [string[]]$to, [string[]]$toExcept = @())

        $root = newSampleRoot $files
        return (@(findFileDependencies $root -From $from -To $to -ToExcept $toExcept) -join ", ")
    }
}

Describe "関数の呼び出し" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "別のパスの関数を呼ぶと、場所・名前・定義元を返す"
            files = @{ "upper/a.ps1" = 'function callA { lowerFn }'; "lower/b.ps1" = 'function lowerFn { 1 }' }
            from = "upper/*"; to = "lower/*"
            expected = "upper/a.ps1 -> lower/b.ps1（lowerFn）"
        }
        @{ name = "同じパスの関数は返さない"
            files = @{ "upper/a.ps1" = 'function callA { sameFn }'; "upper/c.ps1" = 'function sameFn { 1 }'; "lower/b.ps1" = 'function lowerFn { 1 }' }
            from = "upper/*"; to = "lower/*"
            expected = ""
        }
        @{ name = "組み込みのコマンドは返さない"
            files = @{ "upper/a.ps1" = 'function callA { Get-Item . | Out-Null }'; "lower/b.ps1" = 'function lowerFn { 1 }' }
            from = "upper/*"; to = "lower/*"
            expected = ""
        }
        @{ name = "コメントと文字列の中の名前は返さない"
            files = @{ "upper/a.ps1" = "# lowerFn を呼ぶ`nfunction callA { Write-Output 'lowerFn'; Write-Output `"lowerFn`" }"; "lower/b.ps1" = 'function lowerFn { 1 }' }
            from = "upper/*"; to = "lower/*"
            expected = ""
        }
        @{ name = "同じ名前の関数が自分のファイルにもあれば返さない"
            files = @{ "upper/a.ps1" = 'function lowerFn { 2 }; function callA { lowerFn }'; "lower/b.ps1" = 'function lowerFn { 1 }' }
            from = "upper/*"; to = "lower/*"
            expected = ""
        }
        @{ name = "同じ名前の関数が自分のパスの別のファイルにもあれば返さない"
            files = @{ "upper/a.ps1" = 'function callA { lowerFn }'; "upper/c.ps1" = 'function lowerFn { 2 }'; "lower/b.ps1" = 'function lowerFn { 1 }' }
            from = "upper/*"; to = "lower/*"
            expected = ""
        }
        @{ name = "向きを入れ替えると、下が上を呼んだときだけ返す"
            files = @{ "upper/a.ps1" = 'function callA { 1 }'; "lower/b.ps1" = 'function callB { callA }' }
            from = "lower/*"; to = "upper/*"
            expected = "lower/b.ps1 -> upper/a.ps1（callA）"
        }
        @{ name = "向きを入れ替えて、上が下を呼ぶだけなら返さない"
            files = @{ "upper/a.ps1" = 'function callA { lowerFn }'; "lower/b.ps1" = 'function lowerFn { 1 }' }
            from = "lower/*"; to = "upper/*"
            expected = ""
        }
        @{ name = "スクリプトブロックの中の呼び出しも返す"
            files = @{ "upper/a.ps1" = 'function callA { 1..2 | ForEach-Object { lowerFn } }'; "lower/b.ps1" = 'function lowerFn { 1 }' }
            from = "upper/*"; to = "lower/*"
            expected = "upper/a.ps1 -> lower/b.ps1（lowerFn）"
        }
        @{ name = "同じ名前の関数が依存先の 2 つのファイルにあれば、両方を返す"
            files = @{ "upper/a.ps1" = 'function callA { betaFn }'; "lower/b.ps1" = 'function betaFn { 1 }'; "lower/c.ps1" = 'function betaFn { 2 }' }
            from = "upper/*"; to = "lower/*"
            expected = "upper/a.ps1 -> lower/b.ps1（betaFn）, upper/a.ps1 -> lower/c.ps1（betaFn）"
        }
        @{ name = "クラスのメソッドは、関数の定義に数えない"
            files = @{ "upper/a.ps1" = 'function callA { Run }'; "lower/b.ps1" = 'class Worker { [void] Run() { 1 } }' }
            from = "upper/*"; to = "lower/*"
            expected = ""
        }
    ) {
        param ($files, $from, $to, $expected)

        (dependenciesOf $files $from $to) | Should -Be $expected
    }
}

Describe "変数" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "別のパスの最上位の代入を読むと返す"
            files = @{ "upper/a.ps1" = 'function callA { $config }'; "lower/b.ps1" = '$config = 1' }
            expected = 'upper/a.ps1 -> lower/b.ps1（$config）'
        }
        @{ name = "別のパスの関数の中の `$script: の代入を読むと返す"
            files = @{ "upper/a.ps1" = 'function callA { $config }'; "lower/b.ps1" = 'function init { $script:config = 1 }' }
            expected = 'upper/a.ps1 -> lower/b.ps1（$config）'
        }
        @{ name = "`$script: を付けた読みも返す"
            files = @{ "upper/a.ps1" = 'function callA { $script:config }'; "lower/b.ps1" = '$config = 1' }
            expected = 'upper/a.ps1 -> lower/b.ps1（$config）'
        }
        @{ name = "別のパスの関数の中の `$global: の代入を読むと返す"
            files = @{ "upper/a.ps1" = 'function callA { $config }'; "lower/b.ps1" = 'function init { $global:config = 1 }' }
            expected = 'upper/a.ps1 -> lower/b.ps1（$config）'
        }
        @{ name = "別のパスの `$a, `$b = の代入を読むと返す"
            files = @{ "upper/a.ps1" = 'function callA { $second }'; "lower/b.ps1" = '$first, $second = 1, 2' }
            expected = 'upper/a.ps1 -> lower/b.ps1（$second）'
        }
        @{ name = "別のパスの [int]`$a = の代入を読むと返す"
            files = @{ "upper/a.ps1" = 'function callA { $config }'; "lower/b.ps1" = '[int]$config = 1' }
            expected = 'upper/a.ps1 -> lower/b.ps1（$config）'
        }
        @{ name = "型変換が二重の [string][int]`$a = の代入も読む"
            files = @{ "upper/a.ps1" = 'function callA { $config }'; "lower/b.ps1" = '[string][int]$config = 1' }
            expected = 'upper/a.ps1 -> lower/b.ps1（$config）'
        }
        @{ name = "別のパスのクラスのメソッドの中の `$script: の代入を読むと返す"
            files = @{ "upper/a.ps1" = 'function callA { $config }'; "lower/b.ps1" = 'class Worker { [void] Init() { $script:config = 1 } }' }
            expected = 'upper/a.ps1 -> lower/b.ps1（$config）'
        }
        @{ name = "`$env: のようなドライブの変数は返さない"
            files = @{ "upper/a.ps1" = 'function callA { $env:config }'; "lower/b.ps1" = '$env:config = 1' }
            expected = ""
        }
        @{ name = "関数の引数は返さない"
            files = @{ "upper/a.ps1" = 'function callA { param ($config) $config }'; "lower/b.ps1" = '$config = 1' }
            expected = ""
        }
        @{ name = "関数の宣言の引数は返さない"
            files = @{ "upper/a.ps1" = 'function callA ($config) { $config }'; "lower/b.ps1" = '$config = 1' }
            expected = ""
        }
        @{ name = "関数の中の代入は返さない"
            files = @{ "upper/a.ps1" = 'function callA { $config = 2; $config }'; "lower/b.ps1" = '$config = 1' }
            expected = ""
        }
        @{ name = "foreach の変数は返さない"
            files = @{ "upper/a.ps1" = 'function callA { foreach ($config in 1..2) { $config } }'; "lower/b.ps1" = '$config = 1' }
            expected = ""
        }
        @{ name = "スクリプトブロックの引数は返さない"
            files = @{ "upper/a.ps1" = 'function callA { $cb = { param ($config) $config }; & $cb 1 }'; "lower/b.ps1" = '$config = 1' }
            expected = ""
        }
        @{ name = "最上位のスクリプトブロックの引数も返さない"
            files = @{ "upper/a.ps1" = '$cb = { param ($config) $config }'; "lower/b.ps1" = '$config = 1' }
            expected = ""
        }
        @{ name = "自分のパスの最上位にも同じ名前があれば返さない"
            files = @{ "upper/a.ps1" = 'function callA { $config }'; "upper/c.ps1" = '$config = 3'; "lower/b.ps1" = '$config = 1' }
            expected = ""
        }
        @{ name = "別のパスの関数の中だけの代入は、そのパスの変数とみなさない"
            files = @{ "upper/a.ps1" = 'function callA { $config }'; "lower/b.ps1" = 'function init { $config = 1 }' }
            expected = ""
        }
        @{ name = "コメントと文字列の中の名前は返さない"
            files = @{ "upper/a.ps1" = "# `$config を読む`nfunction callA { Write-Output '`$config' }"; "lower/b.ps1" = '$config = 1' }
            expected = ""
        }
    ) {
        param ($files, $expected)

        (dependenciesOf $files "upper/*" "lower/*") | Should -Be $expected
    }
}

Describe "決めた変数と型" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "決めた変数に触ると返す"
            files = @{ "view/a.ps1" = 'function showA { $ui.Title = "a" }' }
            expected = 'view/a.ps1:1 $ui'
        }
        @{ name = "型の式を返す"
            files = @{ "view/a.ps1" = 'function showA { [System.Windows.Media.Brushes]::Red }' }
            expected = "view/a.ps1:1 [System.Windows.Media.Brushes]"
        }
        @{ name = "引数の型を返す"
            files = @{ "view/a.ps1" = 'function showA { param ([System.Windows.Window]$w) $w }' }
            expected = "view/a.ps1:1 [System.Windows.Window]"
        }
        @{ name = "決めていない変数と型は返さない"
            files = @{ "view/a.ps1" = 'function showA { param ([string]$text) $text; [System.IO.Path]::GetTempPath() }' }
            expected = ""
        }
        @{ name = "コメントと文字列の中は返さない"
            files = @{ "view/a.ps1" = "# `$ui と System.Windows.Window`nfunction showA { Write-Output '`$ui System.Windows.Window' }" }
            expected = ""
        }
        @{ name = "対象のパスの外は見ない"
            files = @{ "other/a.ps1" = 'function showA { $window }' }
            expected = ""
        }
    ) {
        param ($files, $expected)

        $arch = newScriptArch (newSampleRoot $files)
        (@(findTypeUses $arch -In "view/*" -Variable ui, window -TypeLike "System.Windows.*") -join ", ") | Should -Be $expected
    }
}

# Set-Variable・New-Variable は、= の代入と同じ条件でそのファイルの最上位の変数として数える
# （関数の外なら数える。関数の中は -Scope が Script・Global・1 以上のときだけ）
Describe "Set-Variable・New-Variable" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "関数の外の -Name 付きを数える"
            code = 'Set-Variable -Name config -Value 1'; counted = $true
        }
        @{ name = "位置引数の名前を数える"
            code = 'Set-Variable config 1'; counted = $true
        }
        @{ name = "別名 sv を数える"
            code = 'sv config 1'; counted = $true
        }
        @{ name = "New-Variable を数える"
            code = 'New-Variable -Name config -Value 1'; counted = $true
        }
        @{ name = "引用符で囲んだ名前を数える"
            code = 'Set-Variable -Name "config" -Value 1'; counted = $true
        }
        @{ name = "-Value を先に書いた位置引数の名前を数える"
            code = 'Set-Variable -Value 1 config'; counted = $true
        }
        @{ name = "-Name を省略した -N も数える"
            code = 'Set-Variable -N config -Value 1'; counted = $true
        }
        @{ name = "-Value に同じ語があっても、-Name でなければ数えない"
            code = 'Set-Variable -Value config -Name other'; counted = $false
        }
        @{ name = "スクリプトブロックの最上位は数える"
            code = '$cb = { Set-Variable -Name config -Value 1 }'; counted = $true
        }
        @{ name = "関数の中の -Scope なしは数えない"
            code = 'function init { Set-Variable -Name config -Value 1 }'; counted = $false
        }
        @{ name = "関数の中の -Scope Script は数える"
            code = 'function init { Set-Variable -Name config -Value 1 -Scope Script }'; counted = $true
        }
        @{ name = "関数の中の -Scope Global は数える"
            code = 'function init { Set-Variable -Name config -Value 1 -Scope Global }'; counted = $true
        }
        @{ name = "関数の中の -Scope 1 は数える"
            code = 'function init { Set-Variable -Name config -Value 1 -Scope 1 }'; counted = $true
        }
        @{ name = "関数の中の -Scope 2 は数える"
            code = 'function init { Set-Variable -Name config -Value 1 -Scope 2 }'; counted = $true
        }
        @{ name = "関数の中の -Scope Local は数えない"
            code = 'function init { Set-Variable -Name config -Value 1 -Scope Local }'; counted = $false
        }
        @{ name = "関数の中の -Scope 0 は数えない"
            code = 'function init { Set-Variable -Name config -Value 1 -Scope 0 }'; counted = $false
        }
        @{ name = "関数の中の -Scope Private は数えない"
            code = 'function init { Set-Variable -Name config -Value 1 -Scope Private }'; counted = $false
        }
        @{ name = "関数の中の Set-Variable の script: 付きの名前は、スコープとして読むので数える"
            code = 'function init { Set-Variable -Name "script:config" -Value 1 }'; counted = $true
        }
        @{ name = "関数の外の Set-Variable の script: 付きの名前は、印を外した名前（config）で数える"
            code = 'Set-Variable -Name "script:config" -Value 1'; counted = $true
        }
        @{ name = "関数の中の Set-Variable の global: 付きの名前は数える"
            code = 'function init { Set-Variable -Name "global:config" -Value 1 }'; counted = $true
        }
        @{ name = "関数の中の Set-Variable の private: 付きの名前は数えない"
            code = 'function init { Set-Variable -Name "private:config" -Value 1 }'; counted = $false
        }
        @{ name = "関数の中の sv の script: 付きの名前は数える"
            code = 'function init { sv "script:config" 1 }'; counted = $true
        }
        @{ name = "関数の中の New-Variable の script: 付きの名前は、名前の一部なので数えない"
            code = 'function init { New-Variable -Name "script:config" -Value 1 }'; counted = $false
        }
        @{ name = "関数の外の New-Variable の script: 付きの名前は、名前そのもの（script:config）なので数えない"
            code = 'New-Variable -Name "script:config" -Value 1'; counted = $false
        }
        @{ name = "関数の中の New-Variable の global: 付きの名前は、名前の一部なので数えない"
            code = 'function init { New-Variable -Name "global:config" -Value 1 }'; counted = $false
        }
        @{ name = "Set-Variable の -Scope と名前の前の印が両方あるとき、印は名前の一部（関数の外は script:config）なので数えない"
            code = 'Set-Variable -Name "script:config" -Value 1 -Scope Script'; counted = $false
        }
        @{ name = "関数の中で -Scope Local と script: 付きの名前が両方あっても、数えない"
            code = 'function init { Set-Variable -Name "script:config" -Value 1 -Scope Local }'; counted = $false
        }
        @{ name = "関数の中で -Scope Script と local: 付きの名前が両方あっても、名前は local:config なので config は数えない"
            code = 'function init { Set-Variable -Name "local:config" -Value 1 -Scope Script }'; counted = $false
        }
        @{ name = "名前を変数で渡す形は、名前が決まらないので数えない"
            code = 'function init { param ($n) Set-Variable -Name $n -Value 1 -Scope Script }'; counted = $false
        }
        @{ name = "別名 nv を数える"
            code = 'nv config 1'; counted = $true
        }
        @{ name = "別名 set を数える"
            code = 'set config 1'; counted = $true
        }
        @{ name = "-Scope を変数で渡す形は、スコープが決まらないので数えない"
            code = 'function init { param ($s) Set-Variable -Name config -Value 1 -Scope $s }'; counted = $false
        }
        @{ name = "関数の中の local: 付きの名前は数えない"
            code = 'function init { Set-Variable -Name "local:config" -Value 1 }'; counted = $false
        }
        @{ name = "値を取る共通の引数の値を、名前と読まない"
            code = 'Set-Variable -ErrorAction SilentlyContinue config 1'; counted = $true
        }
        @{ name = "-OutVariable の値を、名前と読まない"
            code = 'Set-Variable -OutVariable ov config 1'; counted = $true
        }
        @{ name = "モジュールの修飾つきの名前を数える"
            code = 'Microsoft.PowerShell.Utility\Set-Variable config 1'; counted = $true
        }
        @{ name = '関数の外の $variable:名前 = … の代入を数える'
            code = '$variable:config = 1'; counted = $true
        }
        @{ name = '関数の外の ${variable:名前} = … の代入を数える'
            code = '${variable:config} = 1'; counted = $true
        }
        @{ name = '関数の外の ${variable:script:名前} = … は、印を外した名前（config）の代入に数える'
            code = '${variable:script:config} = 1'; counted = $true
        }
        @{ name = '関数の中の $variable:名前 = … は数えない'
            code = 'function init { $variable:config = 1 }'; counted = $false
        }
        @{ name = '関数の中の $variable:script:名前 = … は、$script:名前 = … と同じく数える'
            code = 'function init { $variable:script:config = 1 }'; counted = $true
        }
        @{ name = '関数の中の ${variable:global:名前} = … も、数える'
            code = 'function init { ${variable:global:config} = 1 }'; counted = $true
        }
        @{ name = '関数の中の $variable:local:名前 = … は、その関数のスコープの変数なので数えない'
            code = 'function init { $variable:local:config = 1 }'; counted = $false
        }
        @{ name = '関数の外の $env:名前 = … は数えない'
            code = '$env:config = 1'; counted = $false
        }
    ) {
        param ($code, $counted)

        $files = @{ "upper/a.ps1" = 'function callA { $config }'; "lower/b.ps1" = $code }
        $expected = if ($counted) { 'upper/a.ps1 -> lower/b.ps1（$config）' } else { "" }
        (dependenciesOf $files @("upper/*") @("lower/*")) | Should -Be $expected
    }

    It "Set-Variable の script: 付きの名前は、config の読みと組になる" {
        $files = @{ "upper/a.ps1" = 'function callA { $config }'; "lower/b.ps1" = 'Set-Variable -Name "script:config" -Value 1' }
        (dependenciesOf $files @("upper/*") @("lower/*")) | Should -Be 'upper/a.ps1 -> lower/b.ps1（$config）'
    }

    It "Set-Variable の script: 付きの名前は、Get-Variable の同じ書き方（script:config）とは組にならない" {
        $files = @{ "upper/a.ps1" = 'function callA { Get-Variable -Name "script:config" }'; "lower/b.ps1" = 'Set-Variable -Name "script:config" -Value 1' }
        (dependenciesOf $files @("upper/*") @("lower/*")) | Should -Be ""
    }

    It '関数の中の $variable:名前 = … で作った変数を、その関数の中で読んでも、別のファイルの変数の読みに数えない' {
        $files = @{ "upper/a.ps1" = 'function callA { $variable:config = 2; $config }'; "lower/b.ps1" = '$config = 1' }
        (dependenciesOf $files @("upper/*") @("lower/*")) | Should -Be ""
    }

    It '型を付けた・並べた $variable:名前 の代入は、読みに数えない' -TestCases @(
        @{ code = 'function callA { [int]$variable:config = 1 }' }
        @{ code = 'function callA { $variable:config, $other = 1, 2 }' }
        @{ code = 'function callA { $other, $variable:config = 1, 2 }' }
    ) {
        param ($code)

        $files = @{ "upper/a.ps1" = $code; "lower/b.ps1" = '$config = 1' }
        (dependenciesOf $files @("upper/*") @("lower/*")) | Should -Be ""
    }

    It '代入の右辺の $variable:名前 は、読みに数える' {
        $files = @{ "upper/a.ps1" = 'function callA { $other = $variable:config }'; "lower/b.ps1" = '$config = 1' }
        (dependenciesOf $files @("upper/*") @("lower/*")) | Should -Be 'upper/a.ps1 -> lower/b.ps1（$config）'
    }

    It "Dynamic という名前の変数も、名前の決まらない印と混ざらず数える" {
        $files = @{ "upper/a.ps1" = 'function callA { $Dynamic }'; "lower/b.ps1" = 'Set-Variable -Name Dynamic -Value 1' }
        (dependenciesOf $files @("upper/*") @("lower/*")) | Should -Be 'upper/a.ps1 -> lower/b.ps1（$Dynamic）'
    }
}

# 変数を名前で読む形（Get-Variable・variable: ドライブ）。名前が定数なら、その変数の読みとして数える
Describe "名前で読む変数" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "Get-Variable -Name を数える"
            code = 'function callA { Get-Variable -Name config -ValueOnly }'; counted = $true
        }
        @{ name = "Get-Variable の位置引数の名前を数える"
            code = 'function callA { Get-Variable config -ValueOnly }'; counted = $true
        }
        @{ name = "-ValueOnly を先に書いた位置引数の名前を数える"
            code = 'function callA { Get-Variable -ValueOnly config }'; counted = $true
        }
        @{ name = "-ErrorAction の値があっても名前を数える"
            code = 'function callA { Get-Variable -ErrorAction SilentlyContinue config }'; counted = $true
        }
        @{ name = "別名 gv を数える"
            code = 'function callA { gv config }'; counted = $true
        }
        @{ name = "モジュールの修飾つきの Get-Variable を数える"
            code = 'function callA { Microsoft.PowerShell.Utility\Get-Variable config }'; counted = $true
        }
        @{ name = "script: 付きの名前は、名前そのもの（script:config）なので、config の読みには数えない"
            code = 'function callA { Get-Variable -Name "script:config" }'; counted = $false
        }
        @{ name = 'variable:script:名前 の読みは、印を外した名前（config）の読みに数える'
            code = 'function callA { Get-Item variable:script:config }'; counted = $true
        }
        @{ name = '$variable:script:名前 の読みも、印を外した名前の読みに数える'
            code = 'function callA { $variable:script:config }'; counted = $true
        }
        @{ name = "名前を変数で渡す Get-Variable は、依存に数えない"
            code = 'function callA { param ($n) Get-Variable -Name $n }'; counted = $false
        }
        @{ name = "ワイルドカードの名前は、依存に数えない"
            code = 'function callA { Get-Variable -Name "conf*" }'; counted = $false
        }
        @{ name = '${variable:名前} を数える'
            code = 'function callA { ${variable:config} }'; counted = $true
        }
        @{ name = '$variable:名前 を数える'
            code = 'function callA { $variable:config }'; counted = $true
        }
        @{ name = "Get-Item variable:名前 を数える"
            code = 'function callA { (Get-Item variable:config).Value }'; counted = $true
        }
        @{ name = "引用符つきの Get-Content variable:名前 を数える"
            code = 'function callA { Get-Content "variable:config" }'; counted = $true
        }
        @{ name = "Test-Path variable:名前 を数える"
            code = 'function callA { Test-Path variable:config }'; counted = $true
        }
        @{ name = "variable: の名前が変数のものは、依存に数えない"
            code = 'function callA { param ($n) Get-Item "variable:$n" }'; counted = $false
        }
        @{ name = "variable:\名前 の区切りを、名前に含めず数える"
            code = 'function callA { (Get-Item variable:\config).Value }'; counted = $true
        }
        @{ name = "variable:/名前 の区切りを、名前に含めず数える"
            code = 'function callA { (Get-Item variable:/config).Value }'; counted = $true
        }
        @{ name = "-Path:variable:名前（: でつないだ引数）を数える"
            code = 'function callA { Get-Item -Path:variable:config }'; counted = $true
        }
        @{ name = "-Path variable:名前（空白で分けた引数）を数える"
            code = 'function callA { Get-Item -Path variable:config }'; counted = $true
        }
        @{ name = "-Path:variable:\名前 を数える"
            code = 'function callA { Get-Item -Path:variable:\config }'; counted = $true
        }
        @{ name = "variable: 以外の -Path:ドライブは数えない"
            code = 'function callA { Get-Item -Path:env:config }'; counted = $false
        }
        @{ name = "variable: ドライブへの書き込みは、読みに数えない"
            code = 'function callA { Set-Item variable:config 1 }'; counted = $false
        }
        @{ name = '$env: のようなほかのドライブは数えない'
            code = 'function callA { $env:config; Get-Item env:config }'; counted = $false
        }
    ) {
        param ($code, $counted)

        $files = @{ "upper/a.ps1" = $code; "lower/b.ps1" = '$config = 1' }
        $expected = if ($counted) { 'upper/a.ps1 -> lower/b.ps1（$config）' } else { "" }
        (dependenciesOf $files @("upper/*") @("lower/*")) | Should -Be $expected
    }
}

# 別のファイルで定義したクラスを型として使う依存（名前は [型名]）と、${function:名前} の呼び出し
Describe "クラスの型と function: ドライブ" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "[型]::new() を返す"
            upper = 'function callA { [Worker]::new() }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = "引数の型を返す"
            upper = 'function callA { param ([Worker]$w) $w }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = "キャストを返す"
            upper = 'function callA { param ($y) $x = [Worker]$y }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = "クラスのメソッドの戻り値の型を返す"
            upper = 'class Child { [Worker] Get() { return $null } }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = "継承元の型を返す"
            upper = 'class Child : Worker { }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = "配列の型を返す"
            upper = 'function callA { param ([Worker[]]$w) $w }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = "ジェネリックの引数の型を返す"
            upper = 'function callA { New-Object System.Collections.Generic.List[Worker] }; function callB { [System.Collections.Generic.List[Worker]]::new() }'
            expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = ".NET の型（スクリプトに定義の無い名前）は返さない"
            upper = 'function callA { [System.IO.Path]::GetTempPath(); [void]1 }'; expected = ""
        }
        @{ name = "自分のファイルで定義したクラスは返さない"
            upper = 'class Worker { }; function callA { [Worker]::new() }'; expected = ""
        }
        @{ name = "コメントと文字列の中の型の名前は返さない"
            upper = "# [Worker]`nfunction callA { Write-Output '[Worker]' }"; expected = ""
        }
        @{ name = "`${function:名前} を、その名前の関数の呼び出しとして返す"
            upper = 'function callA { & ${function:betaFn} }'; expected = "upper/a.ps1 -> lower/b.ps1（betaFn）"
        }
        @{ name = "`$function:名前 も返す"
            upper = 'function callA { & $function:betaFn }'; expected = "upper/a.ps1 -> lower/b.ps1（betaFn）"
        }
        @{ name = "New-Object <クラス名> を返す"
            upper = 'function callA { New-Object Worker }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = "New-Object -TypeName <クラス名> を返す"
            upper = 'function callA { New-Object -TypeName Worker }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = "New-Object に変数で渡す型は返さない"
            upper = 'function callA { param ($t) New-Object $t }'; expected = ""
        }
        @{ name = "New-Object の .NET の型は返さない"
            upper = 'function callA { New-Object System.Text.StringBuilder; New-Object byte[] 4 }'; expected = ""
        }
        @{ name = "New-Object -TypeName:<クラス名>（: でつないだ引数）を返す"
            upper = 'function callA { New-Object -TypeName:Worker }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = "New-Object の -TypeName を、ほかの引数のあとに書いても返す"
            upper = 'function callA { New-Object -ArgumentList 1 -TypeName Worker }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = "New-Object の -ArgumentList の値は、型に数えない"
            upper = 'function callA { New-Object -ArgumentList Worker Other }'; expected = ""
        }
        @{ name = "New-Object -ComObject の名前は、型に数えない"
            upper = 'function callA { New-Object -ComObject Worker }'; expected = ""
        }
        @{ name = "New-Object のジェネリックの引数のクラスを返す"
            upper = 'function callA { New-Object ''System.Collections.Generic.List[Worker]'' }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = "New-Object の配列のクラスを返す"
            upper = 'function callA { New-Object ''Worker[]'' 3 }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = "New-Object のジェネリックの引数が複数でも、クラスを返す"
            upper = 'function callA { New-Object ''System.Collections.Generic.Dictionary[string,Worker]'' }'; expected = "upper/a.ps1 -> lower/b.ps1（[Worker]）"
        }
        @{ name = 'Get-Content function:名前 を、関数の呼び出しとしては数えない'
            upper = 'function callA { Get-Content function:betaFn }'; expected = ""
        }
    ) {
        param ($upper, $expected)

        $files = @{ "upper/a.ps1" = $upper; "lower/b.ps1" = "class Worker { }`nfunction betaFn { 1 }" }
        (dependenciesOf $files @("upper/*") @("lower/*")) | Should -Be $expected
    }

    It "別のファイルで定義した enum は、クラスとして数えない" {
        $files = @{ "upper/a.ps1" = 'function callA { [Mode]::A }'; "lower/b.ps1" = 'enum Mode { A; B }' }
        (dependenciesOf $files @("upper/*") @("lower/*")) | Should -Be ""
    }

    It "見る側のパスにも同じ名前のクラスがあれば、そちらを使ったとみなして返さない" {
        $files = @{
            "upper/a.ps1" = 'function callA { [Worker]::new() }'
            "upper/c.ps1" = 'class Worker { }'
            "lower/b.ps1" = 'class Worker { }'
        }
        (dependenciesOf $files @("upper/*") @("lower/*")) | Should -Be ""
    }
}

# 呼ぶ先が構文木から決まらない書き方を、形ごとに返す（呼び出し側が閉じた一覧に載せて確かめる）
Describe "名前を文字列で渡す呼び出し" -Tag Unit {
    It "<name>" -TestCases @(
        @{ name = "Get-Command を返す"
            code = 'function callA { param ($n) (Get-Command $n -CommandType Function).Definition }'; expected = "upper/a.ps1 Get-Command"
        }
        @{ name = "別名 gcm を返す"
            code = 'function callA { param ($n) gcm $n }'; expected = "upper/a.ps1 Get-Command"
        }
        @{ name = "Invoke-Expression を返す"
            code = 'function callA { param ($t) Invoke-Expression $t }'; expected = "upper/a.ps1 Invoke-Expression"
        }
        @{ name = "& に括弧の式を渡す形を返す"
            code = 'function callA { param ($b) & ("a" + $b) }'; expected = "upper/a.ps1 & (式・展開文字列)"
        }
        @{ name = "& に展開文字列を渡す形を返す"
            code = 'function callA { param ($b) & "a$b" }'; expected = "upper/a.ps1 & (式・展開文字列)"
        }
        @{ name = "[scriptblock]::Create を返す"
            code = 'function callA { param ($t) [scriptblock]::Create($t) }'; expected = "upper/a.ps1 [scriptblock]::Create"
        }
        @{ name = '$ExecutionContext.InvokeCommand を返す'
            code = 'function callA { param ($t) $ExecutionContext.InvokeCommand.InvokeScript($t) }'; expected = 'upper/a.ps1 $ExecutionContext.InvokeCommand'
        }
        @{ name = ".AddCommand を返す"
            code = 'function callA { param ($ps) [void]$ps.AddCommand("Get-Item") }'; expected = "upper/a.ps1 .AddCommand"
        }
        @{ name = ".AddScript を返す"
            code = 'function callA { param ($ps) [void]$ps.AddScript("1") }'; expected = "upper/a.ps1 .AddScript"
        }
        @{ name = "function: ドライブの Get-Item を返す"
            code = 'function callA { param ($n) Get-Item "function:$n" }'; expected = "upper/a.ps1 function: ドライブ"
        }
        @{ name = "名前が定数でない Set-Variable を返す"
            code = 'function callA { param ($n) Set-Variable -Name $n -Value 1 }'; expected = "upper/a.ps1 Set-Variable・New-Variable の -Name が定数でない"
        }
        @{ name = "モジュールの修飾つきの Get-Command を返す"
            code = 'function callA { param ($n) Microsoft.PowerShell.Core\Get-Command $n }'; expected = "upper/a.ps1 Get-Command"
        }
        @{ name = "別名 iex を返す"
            code = 'function callA { param ($t) iex $t }'; expected = "upper/a.ps1 Invoke-Expression"
        }
        @{ name = "モジュールの修飾つきの Invoke-Expression を返す"
            code = 'function callA { param ($t) Microsoft.PowerShell.Utility\Invoke-Expression $t }'; expected = "upper/a.ps1 Invoke-Expression"
        }
        @{ name = "Set-Alias を返す"
            code = 'Set-Alias run betaFn'; expected = "upper/a.ps1 Set-Alias・New-Alias"
        }
        @{ name = "New-Alias を返す"
            code = 'New-Alias -Name run -Value betaFn'; expected = "upper/a.ps1 Set-Alias・New-Alias"
        }
        @{ name = "別名 sal を返す"
            code = 'sal run betaFn'; expected = "upper/a.ps1 Set-Alias・New-Alias"
        }
        @{ name = "System.Management.Automation.ScriptBlock の Create を返す"
            code = 'function callA { param ($t) [System.Management.Automation.ScriptBlock]::Create($t) }'; expected = "upper/a.ps1 [scriptblock]::Create"
        }
        @{ name = "Management.Automation.ScriptBlock の Create を返す"
            code = 'function callA { param ($t) [Management.Automation.ScriptBlock]::Create($t) }'; expected = "upper/a.ps1 [scriptblock]::Create"
        }
        @{ name = "Test-Path の function: ドライブを返す"
            code = 'function callA { param ($n) Test-Path "function:$n" }'; expected = "upper/a.ps1 function: ドライブ"
        }
        @{ name = "Get-ChildItem の function: ドライブを返す"
            code = 'function callA { Get-ChildItem function:*Fn }'; expected = "upper/a.ps1 function: ドライブ"
        }
        @{ name = "名前が定数でない Get-Variable を返す"
            code = 'function callA { param ($n) Get-Variable -Name $n -ValueOnly }'; expected = "upper/a.ps1 Get-Variable の名前が定数でない"
        }
        @{ name = "名前の無い Get-Variable を返す"
            code = 'function callA { Get-Variable }'; expected = "upper/a.ps1 Get-Variable の名前が定数でない"
        }
        @{ name = "名前が定数の Get-Variable は返さない（変数の読みとして数える）"
            code = 'function callA { Get-Variable -Name config }'; expected = ""
        }
        @{ name = "名前が定数でない variable: ドライブの読みを返す"
            code = 'function callA { param ($n) Get-Item "variable:$n" }'; expected = "upper/a.ps1 variable: ドライブの名前が定数でない"
        }
        @{ name = "名前が定数の variable: ドライブの読みは返さない"
            code = 'function callA { Get-Item variable:config }'; expected = ""
        }
        @{ name = "別名 nal を返す"
            code = 'nal run betaFn'; expected = "upper/a.ps1 Set-Alias・New-Alias"
        }
        @{ name = "名前を並べた Get-Variable は、名前が決まらないので返す"
            code = 'function callA { Get-Variable a, b }'; expected = "upper/a.ps1 Get-Variable の名前が定数でない"
        }
        @{ name = "パイプで名前を渡す Get-Variable は、名前が決まらないので返す"
            code = 'function callA { "config" | Get-Variable }'; expected = "upper/a.ps1 Get-Variable の名前が定数でない"
        }
        @{ name = "Get-Content の function: ドライブを返す"
            code = 'function callA { Get-Content function:betaFn }'; expected = "upper/a.ps1 function: ドライブ"
        }
        @{ name = "別名 gc・cat・type の function: ドライブを返す"
            code = 'function callA { gc function:betaFn; cat function:betaFn }'; expected = "upper/a.ps1 function: ドライブ, upper/a.ps1 function: ドライブ"
        }
        @{ name = "-Path:function:名前（: でつないだ引数）を返す"
            code = 'function callA { Get-Item -Path:function:betaFn }'; expected = "upper/a.ps1 function: ドライブ"
        }
        @{ name = "Set-Item の variable: ドライブへの書き込みを返す"
            code = 'function callA { Set-Item variable:config 1 }'; expected = "upper/a.ps1 variable: ドライブへの書き込み"
        }
        @{ name = "New-Item の variable: ドライブへの書き込みを返す"
            code = 'function callA { New-Item variable:config -Value 1 }'; expected = "upper/a.ps1 variable: ドライブへの書き込み"
        }
        @{ name = "Remove-Item の variable: ドライブへの書き込みを返す"
            code = 'function callA { Remove-Item variable:config }'; expected = "upper/a.ps1 variable: ドライブへの書き込み"
        }
        @{ name = "Copy-Item の variable: ドライブへの書き込みを返す"
            code = 'function callA { Copy-Item variable:config variable:other }'; expected = "upper/a.ps1 variable: ドライブへの書き込み, upper/a.ps1 variable: ドライブへの書き込み"
        }
        @{ name = "別名 si の variable: ドライブへの書き込みを返す"
            code = 'function callA { si variable:config 1 }'; expected = "upper/a.ps1 variable: ドライブへの書き込み"
        }
        @{ name = "-Path:variable:名前 の書き込みを返す"
            code = 'function callA { Set-Item -Path:variable:config 1 }'; expected = "upper/a.ps1 variable: ドライブへの書き込み"
        }
        @{ name = "variable: 以外のドライブへの書き込みは返さない"
            code = 'function callA { Set-Item env:config 1; Set-Content "$env:TEMP\x" 1 }'; expected = ""
        }
        @{ name = "型を変数で渡す New-Object を返す"
            code = 'function callA { param ($t) New-Object $t }'; expected = "upper/a.ps1 New-Object の型が定数でない"
        }
        @{ name = "型を変数で渡す New-Object -TypeName を返す"
            code = 'function callA { param ($t) New-Object -TypeName $t }'; expected = "upper/a.ps1 New-Object の型が定数でない"
        }
        @{ name = "型が定数の New-Object は返さない"
            code = 'function callA { New-Object System.Text.StringBuilder }'; expected = ""
        }
        @{ name = "New-Object -ComObject は返さない"
            code = 'function callA { New-Object -ComObject Excel.Application }'; expected = ""
        }
        @{ name = "enum は返さない"
            code = 'enum Mode { A; B }'; expected = ""
        }
        @{ name = "& `$変数（呼ぶ値を受け取る形）は返さない"
            code = 'function callA { param ($cb) & $cb 1; . $cb }'; expected = ""
        }
        @{ name = "名前が定数の & は、普通の呼び出しなので返さない"
            code = 'function callA { & "betaFn" }'; expected = ""
        }
        @{ name = "ファイルの読み込み（. `"パス`"）は返さない"
            code = '. "$PSScriptRoot\x.ps1"'; expected = ""
        }
        @{ name = "ふつうの関数・コマンドの呼び出しは返さない"
            code = 'function callA { Get-Item . | Out-Null; betaFn }'; expected = ""
        }
        @{ name = "function: 以外のパスの Get-Item は返さない"
            code = 'function callA { param ($p) Get-Item $p }'; expected = ""
        }
        @{ name = "コメントと文字列の中は返さない"
            code = "# Get-Command `$n`nfunction callA { Write-Output 'Invoke-Expression' }"; expected = ""
        }
    ) {
        param ($code, $expected)

        $arch = newScriptArch (newSampleRoot @{ "upper/a.ps1" = $code; "lower/b.ps1" = 'function betaFn { 1 }' })
        (@(findStringNameCalls $arch -In "upper/*" | ForEach-Object { "$($_.From) $($_.Name)" }) -join ", ") | Should -Be $expected
    }

    It "例外の一覧（From・To・Name）で外せる" {
        $arch = newScriptArch (newSampleRoot @{ "upper/a.ps1" = 'function callA { param ($n) gcm $n }' })
        $items = @(findStringNameCalls $arch)
        $items.Count | Should -Be 1
        $allow = @{ From = "upper/a.ps1"; To = $items[0].To; Name = "Get-Command" }
        @(removeAllowedDependencies $items @($allow)).Count | Should -Be 0
        @(findUnusedAllowances @($allow) $items).Count | Should -Be 0
    }
}

# 同じ意味の書き方の組み合わせ（ドライブの名前の印 × ドライブ × 読む・書く × コマンド・かっこ）を、表にして確かめる
Describe "ドライブの名前の印・かっこ・コマンドの組み合わせ" -Tag Unit {
    # 印（script: など）を外した名前で、変数の読み・関数の呼び出しとして数える（lower/b.ps1 の最上位の変数 $config・関数 betaFn と組になる）
    It "<code> は <expected> に数える" -TestCases @(
        @{ code = 'function callA { Test-Path variable:script:config }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config）' }
        @{ code = 'function callA { Get-Item variable:global:config }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config）' }
        @{ code = 'function callA { Get-Item variable:local:config }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config）' }
        @{ code = 'function callA { Get-ItemProperty variable:config }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config）' }
        @{ code = 'function callA { Get-ItemPropertyValue variable:script:config -Name Value }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config）' }
        @{ code = 'function callA { Resolve-Path variable:config }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config）' }
        @{ code = 'function callA { Test-Path "variable:script:config" }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config）' }
        @{ code = 'function callA { Test-Path ("variable:config") }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config）' }
        @{ code = 'function callA { Get-Item -Path ((''variable:script:config'')) }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config）' }
        @{ code = 'function callA { Get-Item -Path:("variable:config") }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config）' }
        @{ code = 'function callA { Get-Item -Path ("variable:config") }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config）' }
        @{ code = 'function callA { ${variable:global:config} }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config）' }
        @{ code = 'function callA { & ${function:script:betaFn} }'; expected = 'upper/a.ps1 -> lower/b.ps1（betaFn）' }
        @{ code = 'function callA { & $function:global:betaFn }'; expected = 'upper/a.ps1 -> lower/b.ps1（betaFn）' }
        @{ code = 'function callA { Get-Command betaFn -ErrorAction SilentlyContinue }'; expected = 'upper/a.ps1 -> lower/b.ps1（betaFn）' }
        @{ code = 'function callA { Get-Command -CommandType Function betaFn }'; expected = 'upper/a.ps1 -> lower/b.ps1（betaFn）' }
        @{ code = 'function callA { gcm -Name betaFn }'; expected = 'upper/a.ps1 -> lower/b.ps1（betaFn）' }
        @{ code = 'function callA { Get-Command -Name:"betaFn" -All }'; expected = 'upper/a.ps1 -> lower/b.ps1（betaFn）' }
        @{ code = 'function callA { Get-Command $n }'; expected = '' }
        @{ code = 'function callA { Get-Command betaF* }'; expected = '' }
        @{ code = 'function callA { Get-Command -Type Function betaFn }'; expected = '' }
    ) {
        param ($code, $expected)

        $files = @{ "upper/a.ps1" = $code; "lower/b.ps1" = "`$config = 1`nfunction betaFn { 1 }" }
        (dependenciesOf $files @("upper/*") @("lower/*")) | Should -Be $expected
    }

    # 呼ぶ先・変数の名前が決まらない書き方は、形として返す
    It "<code> は、形 <expected> を返す" -TestCases @(
        @{ code = 'function callA { Get-Item (''function:'' + $n) }'; expected = "function: ドライブ" }
        @{ code = 'function callA { Test-Path ("function:$n") }'; expected = "function: ドライブ" }
        @{ code = 'function callA { Get-Item (''variable:'' + $n) }'; expected = "variable: ドライブの名前が定数でない" }
        @{ code = 'function callA { Get-Item -Path ("variable:$n") }'; expected = "variable: ドライブの名前が定数でない" }
        @{ code = 'function callA { Set-Item (''variable:'' + $n) 1 }'; expected = "variable: ドライブへの書き込み" }
        @{ code = 'function callA { Set-Content variable:p 1 }'; expected = "variable: ドライブへの書き込み" }
        @{ code = 'function callA { Set-Content -Path ("variable:p") 1 }'; expected = "variable: ドライブへの書き込み" }
        @{ code = 'function callA { sc variable:p 1 }'; expected = "variable: ドライブへの書き込み" }
        @{ code = 'function callA { Add-Content variable:p 1 }'; expected = "variable: ドライブへの書き込み" }
        @{ code = 'function callA { ac variable:p 1 }'; expected = "variable: ドライブへの書き込み" }
        @{ code = 'function callA { Clear-Content variable:p }'; expected = "variable: ドライブへの書き込み" }
        @{ code = 'function callA { Clear-Item variable:p }'; expected = "variable: ドライブへの書き込み" }
        @{ code = 'function callA { Rename-Item variable:p q }'; expected = "variable: ドライブへの書き込み" }
        @{ code = 'function callA { Move-Item variable:p variable:q }'; expected = "variable: ドライブへの書き込み・variable: ドライブへの書き込み" }
        @{ code = 'function callA { Add-Content function:f "x" }'; expected = "function: ドライブ" }
        @{ code = 'function callA { Set-Content function:f "x" }'; expected = "function: ドライブ" }
        @{ code = 'function callA { Rename-Item function:f g }'; expected = "function: ドライブ" }
        @{ code = 'function callA { Get-ItemProperty function:f }'; expected = "function: ドライブ" }
        @{ code = 'function callA { Get-Command $n }'; expected = "Get-Command" }
        @{ code = 'function callA { Get-Command betaF* }'; expected = "Get-Command" }
        @{ code = 'function callA { Get-Command -Name "betaF?" }'; expected = "Get-Command" }
        @{ code = 'function callA { Get-Command }'; expected = "Get-Command" }
        @{ code = 'function callA { Get-Command -CommandType Function }'; expected = "Get-Command" }
        @{ code = 'function callA { Get-Command betaFn }'; expected = "" }
        @{ code = 'function callA { Get-Command -Type Function betaFn }'; expected = "Get-Command" }
        @{ code = 'function callA { Get-Command -Foo Bar betaFn }'; expected = "Get-Command" }
        @{ code = 'function callA { Set-Variable -Foo Bar -Name x 1 }'; expected = "Set-Variable・New-Variable の -Name が定数でない" }
        @{ code = 'function callA { Set-Content "$env:TEMP\x" 1; Test-Path ($env:TEMP) }'; expected = "" }
        @{ code = 'function callA { Set-Content (Join-Path $dir "x") 1 }'; expected = "" }
    ) {
        param ($code, $expected)

        $arch = newScriptArch (newSampleRoot @{ "upper/a.ps1" = $code; "lower/b.ps1" = 'function betaFn { 1 }' })
        (@(findStringNameCalls $arch -In "upper/*" | ForEach-Object { $_.Name }) -join "・") | Should -Be $expected
    }

    # 名前を文字列の並びで受け取る関数（newWorkerState）は、渡す側が表で教えたときだけ、中の定数の文字列を数える
    Context "名前の並びを受け取る関数" {
        BeforeAll {
            $script:nameListCommands = @{
                workerState = @(@{ Kind = "Function"; Parameter = "functions" }, @{ Kind = "Variable"; Parameter = "variables" })
            }
            function dependenciesWithNameList {
                param ([string]$code, $nameListCommands)

                $files = @{ "upper/a.ps1" = $code; "lower/b.ps1" = "`$config = 1`nfunction betaFn { 1 }`nfunction workerState { 1 }" }
                $arch = newScriptArch (newSampleRoot $files) $nameListCommands
                return (@(findFileDependencies $arch -From @("upper/*") -To @("lower/*")) -join ", ")
            }
        }

        It "<code> は、関数の呼び出しと変数の読みに数える" -TestCases @(
            @{ code = 'function callA { workerState @("betaFn") @("config") }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config・betaFn・workerState）' }
            @{ code = 'function callA { workerState -functions @("betaFn") }'; expected = 'upper/a.ps1 -> lower/b.ps1（betaFn・workerState）' }
            @{ code = 'function callA { workerState -variables @("config") }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config・workerState）' }
            @{ code = 'function callA { workerState -variables @("config") @("betaFn") }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config・betaFn・workerState）' }
            @{ code = 'function callA { $s = workerState @("betaFn", $other) `' + "`n" + '  @("config") }'; expected = 'upper/a.ps1 -> lower/b.ps1（$config・betaFn・workerState）' }
            @{ code = 'function callA { workerState @() @() }'; expected = 'upper/a.ps1 -> lower/b.ps1（workerState）' }
        ) {
            param ($code, $expected)

            dependenciesWithNameList $code $script:nameListCommands | Should -Be $expected
        }

        It "表を渡さなければ、名前の並びの中の文字列は数えない" {
            dependenciesWithNameList 'function callA { workerState @("betaFn") @("config") }' @{} | Should -Be 'upper/a.ps1 -> lower/b.ps1（workerState）'
        }
    }
}

Describe "ファイル単位" -Tag Unit {
    It "関数と変数を、ファイルの組ごとの 1 つにまとめる" {
        $files = @{
            "upper/a.ps1" = 'function callA { betaFn; $alpha }'
            "lower/b.ps1" = '$alpha = 1; function betaFn { 1 }'
        }
        (dependenciesOf $files "upper/*" "lower/*") | Should -Be 'upper/a.ps1 -> lower/b.ps1（$alpha・betaFn）'
    }

    It "ファイルの組が別なら、別の行にする" {
        $files = @{
            "upper/a.ps1" = 'function callA { betaFn; gammaFn }'
            "lower/b.ps1" = 'function betaFn { 1 }'
            "lower/c.ps1" = 'function gammaFn { 1 }'
        }
        (dependenciesOf $files "upper/*" "lower/*") | Should -Be "upper/a.ps1 -> lower/b.ps1（betaFn）, upper/a.ps1 -> lower/c.ps1（gammaFn）"
    }

    It "-ToExcept に当たるファイルは、依存先に数えない" {
        $files = @{
            "upper/a.ps1" = 'function callA { betaFn; gammaFn }'
            "lower/b.ps1" = 'function betaFn { 1 }'
            "lower/skip_view.ps1" = 'function gammaFn { 1 }'
        }
        (dependenciesOf $files "upper/*" "lower/*" "*_view.ps1") | Should -Be "upper/a.ps1 -> lower/b.ps1（betaFn）"
    }

    It "-From に 1 つのファイルだけを指せる" {
        $files = @{
            "upper/a.ps1" = 'function callA { betaFn }'
            "upper/c.ps1" = 'function callC { betaFn }'
            "lower/b.ps1" = 'function betaFn { 1 }'
        }
        (dependenciesOf $files "upper/c.ps1" "lower/*") | Should -Be "upper/c.ps1 -> lower/b.ps1（betaFn）"
    }

    It "解析した結果（newScriptArch）を渡しても、root のパスを渡しても、期待した 1 行になる" {
        $root = newSampleRoot @{ "upper/a.ps1" = 'function callA { betaFn }'; "lower/b.ps1" = 'function betaFn { 1 }' }
        $arch = newScriptArch $root
        $expected = "upper/a.ps1 -> lower/b.ps1（betaFn）"
        (@(findFileDependencies $arch -From "upper/*" -To "lower/*") -join ", ") | Should -Be $expected
        (@(findFileDependencies $root -From "upper/*" -To "lower/*") -join ", ") | Should -Be $expected
    }

    It "構文エラーのあるファイルがあれば、読めないとして失敗する" {
        $root = newSampleRoot @{ "upper/a.ps1" = 'function callA {' }
        { newScriptArch $root } | Should -Throw "*構文エラー*upper/a.ps1*"
    }

    It "別のファイルの型を使うクラスは、構文エラーに数えない" {
        $root = newSampleRoot @{ "upper/a.ps1" = 'class Child : BaseInOtherFile { }' }
        { newScriptArch $root } | Should -Not -Throw
    }
}

Describe "例外" -Tag Unit {
    BeforeAll {
        $script:sampleFiles = @{
            "upper/a.ps1" = 'function callA { betaFn }'
            "mid/m.ps1"   = 'function callM { gammaFn }'
            "lower/b.ps1" = 'function betaFn { 1 }'
            "lower/c.ps1" = 'function gammaFn { 1 }'
        }
        $script:allowAB = @{ From = "upper/a.ps1"; To = "lower/b.ps1"; Name = "betaFn" }
        $script:allowMC = @{ From = "mid/m.ps1"; To = "lower/c.ps1"; Name = "gammaFn" }
    }

    It "例外にした依存は、検出の結果から外れる" {
        $arch = newScriptArch (newSampleRoot $script:sampleFiles)
        $items = @(findFileDependencyItems $arch -From "upper/*" -To "lower/*")
        @(removeAllowedDependencies $items @($script:allowAB)).Count | Should -Be 0
    }

    It "名前が違う例外では外れない" {
        $arch = newScriptArch (newSampleRoot $script:sampleFiles)
        $items = @(findFileDependencyItems $arch -From "upper/*" -To "lower/*")
        $other = @{ From = "upper/a.ps1"; To = "lower/b.ps1"; Name = "otherFn" }
        (formatFileDependencies @(removeAllowedDependencies $items @($other))) | Should -Be "upper/a.ps1 -> lower/b.ps1（betaFn）"
    }

    It "使われなくなった例外は、全部の行の結果の和と突き合わせて決める（別の行で出る依存の例外を、使われないとしない）" {
        $arch = newScriptArch (newSampleRoot $script:sampleFiles)
        $rowUpper = @(findFileDependencyItems $arch -From "upper/*" -To "lower/*")
        $rowMid = @(findFileDependencyItems $arch -From "mid/*" -To "lower/*")
        $allowances = @($script:allowAB, $script:allowMC)

        # 1 行だけと突き合わせると、別の行で出る依存の例外が使われないことになってしまう
        @(findUnusedAllowances $allowances $rowUpper).Count | Should -Be 1
        # 全部の行の和と突き合わせると、どちらも使われている
        @(findUnusedAllowances $allowances ($rowUpper + $rowMid)).Count | Should -Be 0
    }

    It "検出されない依存の例外は、使われなくなった例外として返す" {
        $arch = newScriptArch (newSampleRoot $script:sampleFiles)
        $items = @(findFileDependencyItems $arch -From "upper/*" -To "lower/*")
        $gone = @{ From = "upper/a.ps1"; To = "lower/b.ps1"; Name = "goneFn" }
        $unused = @(findUnusedAllowances @($script:allowAB, $gone) $items)
        $unused.Count | Should -Be 1
        $unused[0].Name | Should -Be "goneFn"
    }
}
