# PSScriptAnalyzer（Microsoft 提供の静的解析）のうち、安全性にかかわるルールだけを有効にした設定。
# 使い方:
#   Invoke-ScriptAnalyzer -Path .\scripts -Recurse -Settings .\tests\meta\PSScriptAnalyzer.security.psd1
# この設定での指摘が 0 件であることを tests\meta\safety.Tests.ps1 で検査する（docs\safety\scans.md「静的解析: PSScriptAnalyzer（Microsoft）」）。
# 書き方・可読性のルール（Write-Host を使わない等）は安全性の判断に関わらないため、ここでは外している。
# 全ルールでの検査は -Settings を付けずに実行する（指摘の内容は docs\safety\scans.md「静的解析: PSScriptAnalyzer（Microsoft）」 の表を参照）。
@{
    IncludeRules = @(
        # 文字列を式として実行しない
        "PSAvoidUsingInvokeExpression",
        # パスワード・資格情報の平文の扱い
        "PSAvoidUsingPlainTextForPassword",
        "PSAvoidUsingConvertToSecureStringWithPlainText",
        "PSUsePSCredentialType",
        "PSAvoidUsingUsernameAndPasswordParams",
        # 通信先の固定・暗号化されない通信
        "PSAvoidUsingComputerNameHardcoded",
        "PSAvoidUsingAllowUnencryptedAuthentication",
        # 危険・非推奨な API
        "PSAvoidUsingBrokenHashAlgorithms",
        "PSAvoidUsingWMICmdlet",
        # 標準コマンドの上書き・別名の定義（意図しない処理に差し替えられるのを防ぐ）
        "PSAvoidOverwritingBuiltInCmdlets",
        "PSAvoidGlobalAliases",
        "PSReservedCmdletChar",
        "PSReservedParams",
        # 文字コード（BOM の無いファイルは読み方によって内容が変わる）
        "PSUseBOMForUnicodeEncodedFile"
    )
}
