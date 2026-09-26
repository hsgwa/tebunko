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
