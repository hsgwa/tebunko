# インデックス名に対する元のフォルダを記録できるかの判断（判断層。ファイルに触らない）。

function testSourceNameRecordable {
    # インデックス名に対する元のフォルダを設定に記録できる名前か。空の名前と、前後に空白（全角空白・NBSP・タブを含む）がある名前は、
    # 設定を読み戻すとき前後の空白が削られて同じ名前にならない（確かめ済みにならない）ため記録しない。
    # 状態層（setIndexSourceFolder）の入口の条件と、画面の文言（記録されないときは毎回確かめる）で同じ判断を使う。
    # 文字数で比べる（-ne はカルチャに従う）
    param (
        [string]$name
    )

    $trimmed = $name.Trim()
    return $trimmed.Length -gt 0 -and $trimmed.Length -eq $name.Length
}
