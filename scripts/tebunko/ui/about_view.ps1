# 「tebunko について」ダイアログ（タブ右端の［⋯］メニュー）に出す文字列の組み立て（判断層）。
# scripts/shared/core/version.ps1 の readVersionFile の結果から、画面に出す版・コミットの文字列を決める。
# 画面に触らないため、そのままテストできる（tests/tebunko/ui/about_view.Tests.ps1）。

function getAboutView {
    # 版・コミットの表示: @{ Version; Commit（出さないときは空） }
    #   versionInfo: readVersionFile の戻り値（$null なら開発中に git から直接起動したときの表示にする）
    param (
        $versionInfo
    )

    if ($null -eq $versionInfo) {
        return @{ Version = "開発版"; Commit = "" }
    }
    return @{ Version = $versionInfo.Tag; Commit = $versionInfo.Sha.Substring(0, 7) }
}
