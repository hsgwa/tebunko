# Support

English | [日本語](SUPPORT.ja.md)

## When you have a problem

1. **Read "[困ったとき](../docs/guide/troubleshooting.md)" (troubleshooting) and "[できないこと](../docs/guide/limitations.md)" (limitations) in the user guide (Japanese).** They list common problems and what to do, the characters that cannot be searched and the files that cannot be read.
2. **Read the logs.** The progress of indexing and the reasons why files could not be read are written to `インデックス作成ログ.txt` in the workspace (by default `%USERPROFILE%\Documents\tebunko_ws`). Errors that stopped indexing are shown on the screen and are also written to the same log. The meaning of each error is described in [docs/design/indexer/errors.md](../docs/design/indexer/errors.md) (Japanese).
3. **Read the design documents.** How the screen, indexing and search work in detail is described in [docs/](../docs/design/index.md) (Japanese).
4. **Search the existing issues.** Check whether the same problem has already been reported in the [issues](https://github.com/hsgwa/tebunko/issues?q=is%3Aissue).

## Reporting

| What | Where |
|---|---|
| It does not work as expected | [Issue (bug)](https://github.com/hsgwa/tebunko/issues/new?template=bug.yml) |
| A request for a new or changed feature | [Issue (feature request)](https://github.com/hsgwa/tebunko/issues/new?template=feature.yml) |
| A security problem | [Private report](https://github.com/hsgwa/tebunko/security/advisories/new) ([SECURITY.md](SECURITY.md)) |

When you paste a screenshot or a log into an issue, first remove paths that contain your user name (`C:\Users\<user name>\...`) and the names of companies and customers. Do not attach the Office files themselves, because they may contain confidential information.

Issues, pull requests and vulnerability reports are welcome in English or Japanese. We reply in English or Japanese.

## Scope of support

tebunko is open source software developed and maintained by an individual. We answer and fix problems as far as we can, but we cannot promise any deadline. The problem may already be fixed in the latest version, so first try the latest version from [Releases](https://github.com/hsgwa/tebunko/releases/latest).
