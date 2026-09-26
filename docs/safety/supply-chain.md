# 供給網（サプライチェーン）とライセンス

第三者の部品を 1 つも使わないため、実行時の供給網に関する論点は次のように閉じる。部品表は [sbom.cdx.json](../../sbom.cdx.json)（CycloneDX 1.6 形式）にある。

| 論点 | 本ツールの状況 |
|---|---|
| 既知の脆弱性（CVE） | **第三者部品に由来する脆弱性は発生しない**。NuGet・npm・PyPI 等のパッケージを使わないため、依存スキャナ（Trivy・Grype・OSV-Scanner 等）を実行しても検出対象が存在しない |
| 部品表（SBOM） | CycloneDX 1.6 形式で、配布 zip と並べてリリースに載せる。構成物は本ツール自身のファイルのみで、`purl`（パッケージ識別子）を持つ部品は 1 件も無い（`safety.Tests.ps1` が確かめる） |
| 前提ソフトウェア | Windows PowerShell 5.1（OS 同梱）、Microsoft Excel（必須）、Word・PowerPoint（旧形式の取り込み時のみ）、.NET Framework 標準アセンブリ。いずれも**同梱せず**、導入先の既存環境を使う |
| 前提ソフトウェアの脆弱性 | 本ツールの管理外であり、導入先の更新管理に従う |
| 本ツールのライセンス | **MIT ライセンス**（[LICENSE](../../LICENSE)）。改変・社内利用・再配布・商用利用ができ、条件は著作権表示とライセンス文の保持のみ。コピーレフト（派生物の公開義務）は無い。無保証（`AS IS`）であることを明記している |
| 第三者ライセンス | 第三者のコードを含まないため、**third-party ライセンスの義務・違反リスクが発生しない**。OSS ライセンス監査（FOSSA・ScanCode 等）の対象となる部品も無い |
| 改ざん検知 | [配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）](scans.md#配布物の完全性カタログハッシュ一覧来歴の署名) のカタログ・ハッシュ一覧・来歴の署名で、配布時点からの変化を受け取り側が検証できる |
| 脆弱性の連絡 | [SECURITY.md](../../.github/SECURITY.ja.md)。公開の Issue ではなく GitHub の非公開の報告窓口（Private vulnerability reporting）で受け付け、7 日以内に受領を返す。修正は最新のマイナー版にパッチ版として出す |

開発・配布の過程は次のように守る（設定はすべてリポジトリにある）。

| 仕組み | 内容 | 設定 |
|---|---|---|
| main への変更 | main へは PR からだけ入れる（直接 push はブランチ保護で禁止）。必須チェック `test`（テスト・安全性の検査・PSScriptAnalyzer・個人情報の検査・`Signed-off-by`）・`pr-title`・`docs`・`analyze`（CodeQL）が通らず、最新の main を取り込んでいない PR はマージできない | [AGENTS.md](../../AGENTS.md)「GitHub の運用」、`.github/workflows/` |
| DCO | コミットに `Signed-off-by` を付ける。コミット時は `commit-msg` フック、PR では CI（`test.yml`）が `tools/check_signoff.ps1` で確かめる | [CONTRIBUTING.md](../../.github/CONTRIBUTING.ja.md) |
| コミット前の検査 | 個人情報・文字コードの検査（`tools/check_commit.ps1`）と、変更に対応するテスト（`tools/run_commit_tests.ps1`）を pre-commit フックで動かす。CI は全ファイル・全履歴を検査する | `tools/hooks/`・`test.yml` |
| GitHub Actions の固定 | ワークフローで使う Actions は、すべてコミットの SHA で版を固定する。ワークフローの権限は既定で読み取りだけにし（`contents: read`。Scorecard は `read-all`）、書き込みが要るジョブだけに足す。checkout は認証情報を残さない（`persist-credentials: false`） | `.github/workflows/*.yml` |
| 依存の更新 | GitHub Actions と、設計書のサイトを作る Python パッケージ（`tools/mkdocs/requirements.txt`。ハッシュで固定し、配布物には入らない）は Dependabot が毎月 PR を出す。PowerShell のモジュール（Pester 5.9.0・PSScriptAnalyzer 1.25.0）は Dependabot の対象外のため、`test.yml` で版を固定する | `.github/dependabot.yml` |
| ワークフローの静的解析 | CodeQL でワークフローを解析する（PR・main への push・毎週）。PowerShell は CodeQL の対象外のため、スクリプトは PSScriptAnalyzer（[静的解析: PSScriptAnalyzer（Microsoft）](scans.md#静的解析-psscriptanalyzermicrosoft)）が受け持つ | `.github/workflows/codeql.yml` |
| OpenSSF Scorecard | ブランチ保護・依存の固定・危険なワークフローの有無などを採点し、結果を公開する（README のバッジ）。採点だけで、マージは止めない | `.github/workflows/scorecard.yml` |
| リリース | タグの push で、テストを通してから配布 zip とインストーラーを作り、カタログ・ハッシュ一覧・SBOM を zip と並べて載せ、来歴に署名する（[配布物の完全性（カタログ・ハッシュ一覧・来歴の署名）](scans.md#配布物の完全性カタログハッシュ一覧来歴の署名)） | `.github/workflows/release.yml` |
