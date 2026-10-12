# 「バージョン情報」ダイアログ（ナビ下の「バージョン情報」）に出す文字列の組み立て（判断層）。
# scripts/shared/core/version.ps1 の readVersionFile の結果から、画面に出す版・コミットの文字列を決める。
# 画面に触らないため、そのままテストできる（tests/tebunko/ui/about_view.Tests.ps1）。

function getAboutView {
    # 版・コミットの表示: @{ Version; Commit（出さないときは空） }
    #   versionInfo: readVersionFile の戻り値（$null なら開発中に git から直接起動したときの表示にする）
    #   devInfo: readDevVersion の戻り値（versionInfo が $null のとき、「開発版」にタグ・コミットを添える。省くと「開発版」だけ）
    param (
        $versionInfo,
        $devInfo = $null
    )

    return getVersionLabel $versionInfo $devInfo
}

${aboutRepositoryUrl} = "https://github.com/hsgwa/tebunko"  # 公開のリポジトリ

function getAboutLinks {
    # ダイアログに出すリンク: @{ Id（ボタンの名前）; Text; Url } の配列。ブラウザで開けるのは、ここに載せた 2 つだけ
    return @(
        @{ Id = "GitHubLink"; Text = "GitHub"; Url = ${aboutRepositoryUrl} }
        @{ Id = "ReleasesLink"; Text = "リリース"; Url = "${aboutRepositoryUrl}/releases" })
}

function testAboutLinkUrl {
    # 開いてよいアドレスか（getAboutLinks に載せたものと、一字一句同じ）
    param (
        [string]$url
    )

    return @(getAboutLinks | Where-Object { $_.Url -ceq $url }).Count -gt 0
}

function getAboutLicenses {
    # ダイアログに出すライセンス: 1 行ずつの文字の配列（tebunko 自身と、同梱するアイコンの形）
    return @(
        "tebunko：MIT ライセンス"
        "アイコンの形（Lucide）：ISC ライセンス")
}
