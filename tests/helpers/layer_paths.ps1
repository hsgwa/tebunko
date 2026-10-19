# 層の決まりのテスト（tests/meta/layers.Tests.ps1）が使う、層に属するファイルの一覧。
# パスは scripts からの相対（区切りは /）で、-like のワイルドカードを使える。BeforeDiscovery と BeforeAll の両方から dot-source する。
#
# 判断層（入力は素の値、出力は素の値。画面にもファイル・COM にも触らない）。
# AGENTS.md「ソースの分け方」と docs/design/structure/source.md の「層」の表の判断層の並びは、この一覧と同じにする（layers.Tests.ps1 が突き合わせる）
$judgmentLayerPaths = @(
    "shared/core/text.ps1"
    "tebunko/index/index_name.ps1"
    "tebunko/core/source_name.ps1"
    "tebunko/index/content_index_format.ps1"
    "tebunko/search/search_query.ps1"
    "tebunko/search/search_gram.ps1"
    "tebunko/index/gram_index.ps1"
    "tebunko/indexer/indexer_decide.ps1"
    "*_view.ps1"
)

# 状態層（ファイル・COM を読み書きする層）のフォルダ。判断層のファイルが、これらのフォルダの中にもある
$stateLayerPaths = @(
    "tebunko/core/*"
    "tebunko/index/*"
    "tebunko/indexer/*"
    "tebunko/search/*"
    "shared/core/*"
    "shared/office/*"
)
