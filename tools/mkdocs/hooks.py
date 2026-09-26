# MkDocs のフック（mkdocs.yml の hooks）。
#
# docs/ の Markdown は GitHub で読むことを前提に書いている。サイトにしたときに GitHub と違う見え方になる所をここで埋める。
#   - docs/ の外へのリンク（../.github/SECURITY.md など）: サイトには含まれないため、GitHub 上のファイルへのリンクに書き換える
import posixpath
import re

# ](../.github/SECURITY.md) や ](../sbom.cdx.json#xxx) の形のリンク
OUTSIDE_LINK = re.compile(r"\]\((\.\./[^)\s#]+)(#[^)\s]*)?\)")


def on_page_markdown(markdown, page, config, files):
    blob = config["repo_url"].rstrip("/") + "/blob/main/"
    page_dir = posixpath.dirname("docs/" + page.file.src_uri)

    def replace(match):
        target = posixpath.normpath(posixpath.join(page_dir, match.group(1)))
        # docs/guide/ などから ../images/ を指すリンクは docs/ の中なので、そのままにする
        if target.startswith("docs/"):
            return match.group(0)
        return "](" + blob + target + (match.group(2) or "") + ")"

    return OUTSIDE_LINK.sub(replace, markdown)
