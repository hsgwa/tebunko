# 場所（シート名・ページ・スライド）を、TSV のファイル名に使える形にする（どのツールからも使う）。
# 抽出（office_extract.ps1・office_reader.ps1 の writeUnits）が書き出す TSV の名前と、それを読む側で同じ規則を使う。
# 関数名の "Index" は、最初に検索のインデックスのために作った名残。比較の一時フォルダの TSV も同じ名前にする

function encodeIndexPlace {
    # インデックスのTSVのファイル名に入れる場所（シート名・ページ・スライド）を符号化する。
    # ファイル名に使えない文字・制御文字と、区切りの _・符号化に使う % を "%XX"（16進数）にする（decodeIndexPlace で元に戻す）。
    #   ・全角に置き換えると `a"b` と `a”b` が同じファイル名になり、後のシートで上書きされるため、元に戻せる形にする
    #   ・場所に _ が残らないため、ファイル名に .xlsx_ 等を含むブックでも、最後の _ でファイル名と場所に分けられる
    param (
        [string]$place
    )

    return [regex]::Replace($place, '[\x00-\x1F"%*/:<>?\\_|]', { param($m) '%{0:X2}' -f [int][char]$m.Value })
}

function decodeIndexPlace {
    # encodeIndexPlace で符号化した場所を元に戻す（encodeIndexPlace が作る "%XX" だけを戻す）
    param (
        [string]$place
    )

    # 符号化した文字が無ければそのまま返す（検索のたびに全TSVの名前を分解するため、置換の呼び出しを省く）
    if ($place.IndexOf("%") -lt 0) {
        return $place
    }
    return [regex]::Replace($place, '%(?:[01][0-9A-F]|2[25AF]|3[ACEF]|5[CF]|7C)', { param($m) [string][char][Convert]::ToInt32($m.Value.Substring(1), 16) })
}

function toIndexFileName {
    # インデックスのTSVのファイル名 "<場所>.tsv" を返す（場所は encodeIndexPlace で符号化する）。
    # 元のファイル名はフォルダ名（= 元のファイル名そのもの）にするため、ファイル名には入れない。
    # ファイル名の上限（255文字）は長いパスの対応（toLongPath）でも超えられないため、超える場合は分かるメッセージで例外にする
    param (
        [string]$place
    )

    $name = "{0}.tsv" -f (encodeIndexPlace $place)
    if ($name.Length -gt ${maxFileNameLength}) {
        throw "インデックスのファイル名が長すぎるため保存できません（$($name.Length) 文字。上限 ${maxFileNameLength} 文字）: ${name}"
    }
    return $name
}
