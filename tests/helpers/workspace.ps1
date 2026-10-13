# テスト用のワークスペース。テストの中で $workspace に入れると、そのテストの間だけ関数が使う場所が変わる
# （関数は既定値で $workspace の場所を使うため）。
#   $workspace = newTestWorkspace @{ IndexDir = "$TestDrive\index" }
#   paths: 差し替える場所（Workspace のプロパティ名 → パス）。dir: ワークスペースのフォルダ

function newTestWorkspace {
    param (
        [hashtable]$paths = @{},
        [string]$dir = "$TestDrive\work"
    )

    $result = [Workspace]::new($dir)
    foreach ($key in $paths.Keys) {
        $result.$key = $paths[$key]
    }
    return $result
}

# ワークスペースの中に置くリンク（ジャンクション・シンボリックリンク）。指す先は、呼ぶ側が必ず $TestDrive の下にする。
# 作れなければ $false を返し（理由は $script:linkError）、呼ぶ側が Set-ItResult -Skipped で飛ばす。
# ジャンクションは管理者の権限が要らないので、CI（GITHUB_ACTIONS がある）で作れなければ飛ばさず例外にする。
# シンボリックリンクは、作れない環境（権限・開発者モードが無い）があるため、CI でも飛ばす

function newTestJunction {
    param (
        [string]$link,
        [string]$target
    )

    [System.IO.Directory]::CreateDirectory($target) | Out-Null
    try {
        New-Item -ItemType Junction -Path $link -Target $target -ErrorAction Stop | Out-Null
        return $true
    } catch {
        $script:linkError = $_.Exception.Message
        if ($env:GITHUB_ACTIONS) {
            throw "ジャンクションを作れませんでした（CI では飛ばさず失敗にする）：$($script:linkError)"
        }
        return $false
    }
}

function newTestSymlink {
    param (
        [string]$link,
        [string]$target,
        [switch]$file   # 指す先がファイルのとき
    )

    if (!$file) {
        [System.IO.Directory]::CreateDirectory($target) | Out-Null
    }
    try {
        New-Item -ItemType SymbolicLink -Path $link -Target $target -ErrorAction Stop | Out-Null
        return $true
    } catch {
        $script:linkError = $_.Exception.Message
        return $false
    }
}
