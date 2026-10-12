# 表（DataGrid）に外枠の線を付けないことのテスト。
# 外枠はデザインの見本に無い余計な線なので、共通のスタイルでも、各画面の表でも付けない。
# XAML を XML として読み、TargetType が DataGrid のスタイルの BorderThickness と、DataGrid 要素の BorderThickness を見る（画面は開かない）。
BeforeAll {
    $script:xaml = "http://schemas.microsoft.com/winfx/2006/xaml"
    $script:docs = New-Object System.Collections.Generic.List[object]
    foreach ($f in Get-ChildItem -LiteralPath (Resolve-Path "$PSScriptRoot\..\..\scripts").Path -Recurse -Filter "*.xaml") {
        $doc = New-Object System.Xml.XmlDocument
        $doc.Load($f.FullName)
        $script:docs.Add(@{ Name = $f.Name; Doc = $doc })
    }
    function script:isZero($value) {
        $null -ne $value -and $value.Trim() -match '^0(\s*,\s*0){0,3}$'
    }
}

Describe "表に外枠を付けない" -Tag Meta {
    It "TargetType が DataGrid のスタイルは、BorderThickness を 0 にしている（BasedOn で既定を継ぐ場合を除く）" {
        $bad = New-Object System.Collections.Generic.List[string]
        $found = 0
        foreach ($d in $script:docs) {
            $styles = $d.Doc.SelectNodes("//*[local-name()='Style'][@TargetType='DataGrid' or @TargetType='{x:Type DataGrid}']")
            foreach ($style in $styles) {
                $found++
                $setter = $style.SelectSingleNode("*[local-name()='Setter'][@Property='BorderThickness']")
                if ($null -eq $setter) {
                    if (-not $style.HasAttribute('BasedOn')) { $bad.Add("$($d.Name)：BorderThickness の Setter が無い") }
                }
                elseif (-not (script:isZero $setter.GetAttribute('Value'))) {
                    $bad.Add("$($d.Name)：BorderThickness の値が 0 でない（$($setter.GetAttribute('Value'))）")
                }
            }
        }
        $found | Should -BeGreaterThan 0
        ($bad -join "`n") | Should -Be ""
    }

    It "DataGrid 要素は、0 以外の BorderThickness を付けない（属性でも要素でも）" {
        $bad = New-Object System.Collections.Generic.List[string]
        foreach ($d in $script:docs) {
            foreach ($grid in $d.Doc.SelectNodes("//*[local-name()='DataGrid']")) {
                if ($grid.HasAttribute('BorderThickness') -and -not (script:isZero $grid.GetAttribute('BorderThickness'))) {
                    $bad.Add("$($d.Name)：$($grid.GetAttribute('Name', $script:xaml))")
                }
            }
            foreach ($element in $d.Doc.SelectNodes("//*[local-name()='DataGrid.BorderThickness']")) {
                if (-not (script:isZero $element.InnerText)) { $bad.Add("$($d.Name)：DataGrid.BorderThickness の要素") }
            }
        }
        ($bad -join "`n") | Should -Be ""
    }

    It "インデックス管理の一覧は、表の上辺だけに線を引く（表を包む Border の BorderThickness が 0,1,0,0）" {
        $list = $script:docs | Where-Object { $_.Name -eq "index_list.xaml" }
        $grid = $list.Doc.SelectSingleNode("//*[local-name()='DataGrid'][@*[local-name()='Name']='IndexGrid']")
        $grid | Should -Not -BeNullOrEmpty
        $border = $grid.ParentNode
        while ($null -ne $border -and $border.LocalName -ne 'Border') { $border = $border.ParentNode }
        $border | Should -Not -BeNullOrEmpty
        $border.GetAttribute('BorderThickness') -replace '\s', '' | Should -Be "0,1,0,0"
    }
}
