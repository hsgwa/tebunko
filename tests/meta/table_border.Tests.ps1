# 表（DataGrid）に外枠の線を付けないことのテスト。
# 外枠は Figma の見本に無い余計な線なので、共通のスタイルでも、各画面の表でも付けない。
# XAML を XML として読み、TargetType が DataGrid のスタイルの BorderThickness と、DataGrid 要素の BorderThickness を見る（画面は開かない）。
BeforeAll {
    $script:xamlFiles = Get-ChildItem -LiteralPath (Resolve-Path "$PSScriptRoot\..\..\scripts").Path -Recurse -Filter "*.xaml"
    function script:loadXaml($file) {
        $doc = New-Object System.Xml.XmlDocument
        $doc.Load($file.FullName)
        $doc
    }
    function script:isZero($value) {
        $null -ne $value -and $value.Trim() -match '^0(\s*,\s*0){0,3}$'
    }
}

Describe "表に外枠を付けない" -Tag Meta {
    It "TargetType が DataGrid のスタイルは、BorderThickness を 0 にしている" {
        $bad = New-Object System.Collections.Generic.List[string]
        $found = 0
        foreach ($f in $script:xamlFiles) {
            $doc = script:loadXaml $f
            foreach ($style in $doc.SelectNodes("//*[local-name()='Style'][@TargetType='DataGrid']")) {
                $found++
                $setter = $style.SelectSingleNode("*[local-name()='Setter'][@Property='BorderThickness']")
                if ($null -eq $setter -or -not (script:isZero $setter.GetAttribute('Value'))) { $bad.Add($f.Name) }
            }
        }
        $found | Should -BeGreaterThan 0
        ($bad -join "`n") | Should -Be ""
    }

    It "DataGrid 要素は、0 以外の BorderThickness を付けない" {
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($f in $script:xamlFiles) {
            $doc = script:loadXaml $f
            foreach ($grid in $doc.SelectNodes("//*[local-name()='DataGrid']")) {
                if ($grid.HasAttribute('BorderThickness') -and -not (script:isZero $grid.GetAttribute('BorderThickness'))) {
                    $bad.Add("$($f.Name)：$($grid.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml'))")
                }
            }
        }
        ($bad -join "`n") | Should -Be ""
    }
}
