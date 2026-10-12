# 設定ファイルの場所の欄（SettingsFileText）が、枠のない専用の見た目であることのテスト。
# 共通の入力欄の見た目（枠・高さ 30）を使うと、枠が出て、折り返した 2 行目が切れる。画面の自動操作では枠が見えないので、XAML を XML として読んで確かめる。
BeforeAll {
    $script:doc = New-Object System.Xml.XmlDocument
    $script:doc.Load((Resolve-Path "$PSScriptRoot\..\..\scripts\tebunko\xaml\settings\settings.xaml").Path)
    $script:xamlNs = "http://schemas.microsoft.com/winfx/2006/xaml"
}

Describe "設定ファイルの場所の欄" -Tag Meta {
    BeforeAll {
        $script:box = $script:doc.SelectSingleNode("//*[local-name()='TextBox'][@*[local-name()='Name']='SettingsFileText']")
        $script:styleKey = if ($script:box) { ($script:box.GetAttribute("Style") -replace '^\{StaticResource\s+(\S+)\}$', '$1') } else { "" }
        $script:style = $script:doc.SelectNodes("//*[local-name()='Style']") | Where-Object { $_.GetAttribute("Key", $script:xamlNs) -eq $script:styleKey } | Select-Object -First 1
    }

    It "専用の Style を使っている" {
        $script:box | Should -Not -BeNullOrEmpty
        $script:styleKey | Should -Not -BeNullOrEmpty
        $script:style | Should -Not -BeNullOrEmpty
    }

    It "欄そのものには、高さと枠の太さを書かない（書くと専用の Style の指定を上書きする）" {
        $script:box.HasAttribute("Height") | Should -BeFalse
        $script:box.HasAttribute("BorderThickness") | Should -BeFalse
        $script:box.SelectNodes("*[contains(local-name(), 'TextBox.')]") | Where-Object { $_.LocalName -in @("TextBox.Height", "TextBox.BorderThickness") } | Should -BeNullOrEmpty
    }

    It "その Style は、枠の太さが 0 で、高さを決めず、見た目の中に枠（Border）を持たない" {
        $setter = { param ($name) $script:style.SelectSingleNode("*[local-name()='Setter'][@Property='$name']") }
        (& $setter "BorderThickness").GetAttribute("Value") | Should -Be "0"
        & $setter "Height" | Should -BeNullOrEmpty
        $script:style.SelectNodes(".//*[local-name()='Border']").Count | Should -Be 0
        $script:style.HasAttribute("BasedOn") | Should -BeFalse -Because "共通の入力欄の見た目（枠・高さ 30）を引き継がない"
    }
}