# PSScriptAnalyzer の全体の検査の設定（scripts\ と tools\ 用）。tools\check_analyzer.ps1 が使う。
# 既定の全規則（Error・Warning・Information）に、既定で無効の次の規則を足す。
#   PSUseCompatibleSyntax / PSUseCompatibleCommands … Windows PowerShell 5.1 に無い構文・コマンドを止める
#   PSAvoidSemicolonsAsLineTerminators / PSPlaceOpenBrace … 書き方の誤りを拾う
# 次の規則は、指摘のほとんどが誤りでなく、新しいコードで正しく増えるため除く（docs\safety\scans.md の表と同じ）。
# 指摘が残る規則は tests\meta\PSScriptAnalyzer.baseline.tsv にファイルごとの件数で記録し、増えたら失敗にする。
# PSUseConstrainedLanguageMode は足さない（tests\meta\safety.Tests.ps1 が受け持つ）。
@{
    IncludeDefaultRules = $true
    ExcludeRules = @(
        # COM の呼び出し（Workbooks.Open など）は位置で渡すしかない
        "PSAvoidUsingPositionalParameters",
        # WPF のイベントハンドラーの引数（sender・e）
        "PSReviewUnusedParameter",
        # WPF のハンドラーの引数名 sender
        "PSAvoidAssignmentToAutomaticVariable",
        # -ErrorVariable・${...} など別の構文で使う変数を見られない
        "PSUseDeclaredVarsMoreThanAssignments",
        # 継承元の型が別ファイルにある（読み込む順で解決することは structure.Tests.ps1 が確かめる）
        "TypeNotFound"
    )
    Rules = @{
        PSUseCompatibleSyntax = @{
            Enable         = $true
            TargetVersions = @("5.1")
        }
        PSUseCompatibleCommands = @{
            Enable         = $true
            TargetProfiles = @("win-48_x64_10.0.17763.0_5.1.17763.316_x64_4.0.30319.42000_framework")
        }
        PSAvoidSemicolonsAsLineTerminators = @{ Enable = $true }
        PSPlaceOpenBrace = @{
            Enable             = $true
            OnSameLine         = $true
            NewLineAfter       = $true
            IgnoreOneLineBlock = $true
        }
    }
}
