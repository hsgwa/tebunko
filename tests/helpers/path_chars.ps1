# フォルダ名に入れて起動を確かめる文字の表（tests/meta/launcher.Tests.ps1・tests/meta/installer.Tests.ps1 が
# BeforeDiscovery で dot-source し、-TestCases に渡す）。
#   label はテスト名に出す説明、ch はフォルダ名に入れる文字（PowerShell・cmd の構文で特別な意味を持つもの・日本語など）

function getPathCharCases {
    return @(
        @{ label = "一重引用符 '"; ch = "'" }
        @{ label = "右の一重引用符 U+2019"; ch = [string][char]0x2019 }
        @{ label = "空白"; ch = " " }
        @{ label = "丸括弧"; ch = "(1)" }
        @{ label = "&"; ch = "&" }
        @{ label = "^"; ch = "^" }
        @{ label = ";"; ch = ";" }
        @{ label = ","; ch = "," }
        @{ label = "="; ch = "=" }
        @{ label = "単独の %"; ch = "50%" }
        @{ label = "!"; ch = "!" }
        @{ label = "$"; ch = '$' }
        @{ label = "バッククォート"; ch = '`' }
        @{ label = "#"; ch = "#" }
        @{ label = "@"; ch = "@" }
        @{ label = "日本語"; ch = "営業部" }
        @{ label = "角括弧"; ch = "[1]" }
        @{ label = "角括弧の中が空 []"; ch = "[]" }
        @{ label = "開く角括弧だけ"; ch = "[" }
        @{ label = "閉じる角括弧だけ"; ch = "]" }
        @{ label = "引用符・&・角括弧の組み合わせ"; ch = "O'aa&[]fds" }
        @{ label = "バッククォートと開く角括弧"; ch = '`[' }
        @{ label = "バッククォート 2 つ"; ch = '``' }
        @{ label = "バッククォートと角括弧の組み合わせ"; ch = '`a[1]`b' }
    )
}
