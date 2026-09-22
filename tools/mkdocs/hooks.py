# MkDocs のフック（mkdocs.yml の hooks）。
#
# docs/ の Markdown は GitHub で読むことを前提に書いている。サイトにしたときに GitHub と違う見え方になる所をここで埋める。
#   - 先頭ページ: docs/ に index.md は無く、00_index.md が目次にあたる。これをサイトの先頭（/）に置く
#   - docs/ の外へのリンク（../.github/SECURITY.md など）: サイトには含まれないため、GitHub 上のファイルへのリンクに書き換える
import posixpath
import re

HOME_PAGE = "00_index.md"

# ](../.github/SECURITY.md) や ](../sbom.cdx.json#xxx) の形のリンク
OUTSIDE_LINK = re.compile(r"\]\((\.\./[^)\s#]+)(#[^)\s]*)?\)")


def on_files(files, config):
    home = files.get_file_from_path(HOME_PAGE)
    if home is not None:
        home.dest_uri = "index.html"
        home.url = "./"
    return files


def on_page_markdown(markdown, page, config, files):
    blob = config["repo_url"].rstrip("/") + "/blob/main/"
    page_dir = posixpath.dirname("docs/" + page.file.src_uri)

    def replace(match):
        target = posixpath.normpath(posixpath.join(page_dir, match.group(1)))
        return "](" + blob + target + (match.group(2) or "") + ")"

    return OUTSIDE_LINK.sub(replace, markdown)
