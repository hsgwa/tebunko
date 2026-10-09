# データの置き場所（shared\core\data_dir.ps1）のテスト
BeforeAll {
    . "$PSScriptRoot\..\..\helpers\load.ps1"

    function denyCreateFiles {
        # フォルダに、今の利用者がファイルを作れないようにする（Program Files に置いたときと同じく書き込めない状態）。戻すときの規則を返す
        param ([string]$dir)
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            [System.Security.Principal.WindowsIdentity]::GetCurrent().User, "CreateFiles", "Deny")
        $acl = Get-Acl -LiteralPath $dir
        $acl.AddAccessRule($rule)
        Set-Acl -LiteralPath $dir -AclObject $acl
        return $rule
    }

    function restoreAccess {
        param ([string]$dir, $rule)
        $acl = Get-Acl -LiteralPath $dir
        [void]$acl.RemoveAccessRule($rule)
        Set-Acl -LiteralPath $dir -AclObject $acl
    }
}

Describe "testWritableFolder" -Tag Io {
    It "ファイルを作れるフォルダなら `$true で、試しに作ったファイルは残さない" {
        $dir = "$TestDrive\書ける"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        testWritableFolder $dir | Should -Be $true
        @(Get-ChildItem -LiteralPath $dir -Force).Count | Should -Be 0
    }

    It "フォルダが無い・空なら `$false" {
        testWritableFolder "$TestDrive\無いフォルダ" | Should -Be $false
        testWritableFolder "" | Should -Be $false
    }

    It "ファイルを作れないフォルダなら `$false" {
        $dir = "$TestDrive\書けない"
        [System.IO.Directory]::CreateDirectory($dir) | Out-Null
        $rule = denyCreateFiles $dir
        try {
            testWritableFolder $dir | Should -Be $false
        } finally {
            restoreAccess $dir $rule
        }
    }
}

Describe "getDataDir" -Tag Io {
    It "<Case>" -TestCases @(
        @{ Case = "ツールのフォルダに書き込めれば、そのフォルダ"; Writable = $true }
        @{ Case = "書き込めない（存在しない）フォルダなら、渡した逃げ先"; Writable = $false }
    ) {
        $fallback = "$TestDrive\逃げ先"
        $root = "$TestDrive\ツール-$Writable"
        if ($Writable) {
            [System.IO.Directory]::CreateDirectory($root) | Out-Null
        }
        getDataDir $root $fallback | Should -Be $(if ($Writable) { $root } else { $fallback })
    }

    It "読み取り専用のフォルダなら、渡した逃げ先" {
        $root = "$TestDrive\読み取り専用のツール"
        [System.IO.Directory]::CreateDirectory($root) | Out-Null
        $rule = denyCreateFiles $root
        try {
            $dir = getDataDir $root "$TestDrive\逃げ先"
        } finally {
            restoreAccess $root $rule
        }
        $dir | Should -Be "$TestDrive\逃げ先"
    }
}