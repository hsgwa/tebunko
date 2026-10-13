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
