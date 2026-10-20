# PSScriptAnalyzer の全体の検査の設定（tests\ 用）。tools\check_analyzer.ps1 が使う。
# PSScriptAnalyzer.psd1 と同じだが、次が違う。
#   - PSUseCompatibleCommands を入れない（Pester の Should・Describe などがプロファイルに無く、誤検知になる）
#   - PSAvoidUsingWriteHost（テストの補助が進み具合を出す）・PSAvoidGlobalVars（Pester のモックが $global: で状態を渡す）も除く
@{
    IncludeDefaultRules = $true
    ExcludeRules = @(
        "PSAvoidUsingPositionalParameters",
        "PSReviewUnusedParameter",
        "PSAvoidAssignmentToAutomaticVariable",
        "PSUseDeclaredVarsMoreThanAssignments",
        "TypeNotFound",
        "PSAvoidUsingWriteHost",
        "PSAvoidGlobalVars"
    )
    Rules = @{
        PSUseCompatibleSyntax = @{
            Enable         = $true
            TargetVersions = @("5.1")
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
