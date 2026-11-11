# 動作確認用のサンプルを作る tools\new_sample.ps1 のテスト
BeforeAll {
    $rootDir = (Resolve-Path "$PSScriptRoot\..\..").Path
    $newSample = "$rootDir\tools\new_sample.ps1"
}

Describe "new_sample.ps1 -PlanOnly" -Tag Unit {
    It "-PlanOnly を同じ引数で 2 回流すと同じ一覧が返る" {
        $first = @(& $newSample -PlanOnly -Count 40)
        $second = @(& $newSample -PlanOnly -Count 40)
        ($first -join "`n") | Should -Be ($second -join "`n")
    }

    It "-PlanOnly の件数が -Count と一致し、xlsx・docx・pptx の 3 種がすべて含まれる" {
        $plan = @(& $newSample -PlanOnly -Count 40)
        $plan.Count | Should -Be 40
        $extensions = @($plan | ForEach-Object { [System.IO.Path]::GetExtension($_) } | Sort-Object -Unique)
        $extensions | Should -Be @(".docx", ".pptx", ".xlsx")
    }

    It "-PlanOnly の一覧のパスが部署フォルダ配下で、実在の組織名を含まない" {
        $plan = @(& $newSample -PlanOnly -Count 40)
        foreach ($rel in $plan) {
            $rel | Should -Match '^\(株\)山田商事\\[^\\]+\\\d{4}年度\\[^\\]+$'
        }
    }

    It "上限の 500 件でも名前が重ならない" {
        $plan = @(& $newSample -PlanOnly -Count 500)
        $plan.Count | Should -Be 500
        @($plan | Select-Object -Unique).Count | Should -Be 500
    }

    It "-Count が 500 を超えると失敗する" {
        { & $newSample -PlanOnly -Count 501 } | Should -Throw
    }
}

Describe "new_sample.ps1 の生成" -Tag Office {
    BeforeAll {
        $outDir = Join-Path $TestDrive "sample"
        $plan = @(& $newSample -PlanOnly -Count 6)
        & $newSample -OutDir $outDir -Count 6 | Out-Null
    }

    It "-OutDir に作ると一覧どおりのファイルができる" {
        $made = @(Get-ChildItem -LiteralPath $outDir -Recurse -File | ForEach-Object { $_.FullName.Substring($outDir.Length + 1) })
        ($made | Sort-Object) -join "`n" | Should -Be (($plan | Sort-Object) -join "`n")
    }

    It "作ったファイルに作成者名が残らない" {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        foreach ($rel in $plan) {
            $zip = [System.IO.Compression.ZipFile]::OpenRead((Join-Path $outDir $rel))
            try {
                $entry = $zip.GetEntry("docProps/core.xml")
                $reader = New-Object System.IO.StreamReader($entry.Open())
                $xml = $reader.ReadToEnd()
                $reader.Close()
            }
            finally { $zip.Dispose() }
            if ($xml -match '<dc:creator>([^<]*)</dc:creator>') { $Matches[1] | Should -Be "test" }
        }
    }
}
