# メンテナの運用

メンテナ（リポジトリの持ち主）とそのエージェントが、作業を進めてマージ・リリースするまでの決まり。外部からの貢献の手順は [CONTRIBUTING.md](CONTRIBUTING.md)、コードを書くときの決まりは [AGENTS.md](../AGENTS.md) にある。

## 作業場所

本リポジトリ配下での作業は git worktree（EnterWorktree 等）で隔離したツリー上で行い、作業ツリーを直接編集しない。

ブランチ（worktree）を作るときは、先に GitHub の main を取得し、その最新から作る。手元の `master` や古いブランチを起点にしない。

```
git fetch origin
git worktree add -b worktree-<名前> .claude/worktrees/<名前> origin/main
```

一つのセッションを続けて別の機能に取りかかるときは、今の worktree で続けるか、別の worktree を新しく作るかを利用者に確認してから始める。別の worktree にする場合も、上と同じく最新の `origin/main` から作る。

作業内容をメインブランチへマージしたら、その worktree は削除する。マージ済みで不要になった worktree を残さない。

squash merge では手元のコミットが main に入らないため、`git branch -d` はブランチを消せない。PR がマージ済み（`MERGED`）であることを確かめてから `-D` で消す。

```
gh pr view worktree-<名前> --json state --jq .state
git worktree remove .claude/worktrees/<名前>
git branch -D worktree-<名前>
```

未コミットの変更が残っている worktree は削除しない。コミットするか破棄するかを利用者に確認してから削除する。

## 作業の管理（GitHub Projects）

作業は非公開の GitHub Project「tebunko」（持ち主 `hsgwa`、リポジトリにつないである）で管理する。Issue の一覧を作業の管理に使わない。Project の操作は `gh project` で行う（トークンに `project` スコープが要る）。

- **作業は Draft アイテムで始める。** やることは Project の Draft アイテム（Issue にしない下書き）に書き、Issue は立てない。目的・やること・終わりの条件は Draft の本文に書き、途中で分かったことも書き足す。
- **Issue にするのは、内容がまとまってからにする。** 目的と範囲が固まり、公開の記録に残す価値があると利用者が判断したときに、Draft を Issue に変える（Project の画面の「Convert to issue」）。Issue に変えるかどうかは利用者が決める。エージェントは変えたほうがよいと考えたら提案する。
- **外部から来た Issue はそのまま受ける。** 不具合の報告・機能の要望はこれまでどおり Issue のテンプレート（`.github/ISSUE_TEMPLATE/`）で受け、Project に取り込む。安全性の問題は [SECURITY.md](SECURITY.md) の非公開の窓口で受ける。
- **Status で進み具合を表す。**

  | Status | 意味 | 動かす人 |
  |---|---|---|
  | `Backlog` | 思いついたこと。まだ手を付けない | 利用者・エージェント |
  | `Ready` | 手を付けてよい。上にあるものほど先にやる | 利用者 |
  | `In progress` | エージェントが作業中 | エージェント |
  | `Review` | PR を出し、利用者の確認を待っている | エージェント |
  | `Done` | マージした（PR がマージされると自動で移る） | 自動 |

- **エージェントは `Ready` の上から取る。** 取ったら `In progress` にし、worktree を作って進める。PR を出したら、PR を Project に足してアイテムと並べ、アイテムを `Review` にする。利用者から直接頼まれた作業も、Draft アイテムを作ってから始める。
- **1 つのアイテムは 1 つの目的にし、PR 1 本で片付ける。** 作業を細かく分けてアイテムや PR を並べない。1 本では大きすぎるときは、分け方を利用者に示し、了承を得てから分ける。ci・docs の小さな直しは 1 本ずつ出さず、「運用の改善」のような 1 つのアイテムにまとめて PR 1 本にする。

## マージ

**エージェントは自分の判断でマージしない。** 次の条件をそろえた PR を Project の `Review` に並べ、利用者にまとめて提案し、許可を得てからマージする。PR 1 本ごとに都度たずねず、`Review` にたまったものを一度に諮る。

1. 最新の main を取り込んである（[AGENTS.md](../AGENTS.md) の「GitHub の運用」の `git merge origin/main`）。
2. CI がすべて通っている（`gh pr checks <番号>`）。失敗・実行中のものが 1 つでもあればマージしない。
3. PR テンプレートの「確認したこと」をすべて済ませた。画面を変えた場合は、実際に起動して見た。
4. 差分を読み直し、PR の目的と関係のない変更・個人情報が入っていない。
5. ラベルが付いている。

提案には次を書く。

- `Review` にある PR の一覧（番号とタイトル。タイトルは main のコミットのタイトルになる）と、それぞれの変更の要点
- CI の結果と、条件 1〜5 を満たしていること（満たしていない PR は一覧に入れず、理由を添える）
- マージする順番
- マージの後に出すリリース（版の番号、または出さない理由。下の「リリース」）
- `breaking` の PR は、移行の手順

利用者は一覧をまとめて許可してよい（「全部」「#80 以外」など）。許可したものだけを、示した順に次でマージする。

```
gh pr merge <番号> --squash
```

- `--admin` でブランチ保護を飛ばさない。
- Dependabot の PR も同じ条件・同じ手順で諮る。
- マージしたら、上の「作業場所」の手順で worktree とブランチを消す。続けて下の「リリース」の要否を判断する。

## リリース

`v<メジャー>.<マイナー>.<パッチ>`（SemVer）のタグを main のコミットに付けて push すると、`release.yml` が配布 zip を GitHub Release に載せる。

**今は 0.x（`v0.<マイナー>.<パッチ>`）を続ける。** 1.0.0 に上げるのは利用者が決める。0.x の間、上げる桁は前のタグ以降にマージされた PR のラベルで次のように決める。上から順に見て、最初に当てはまったものにする。

| マージされた PR のラベル | 0.x の間 | 1.0.0 以降 |
|---|---|---|
| `breaking` が 1 つでもある | マイナー（0.3.1 → 0.4.0） | メジャー（1.3.1 → 2.0.0） |
| `enhancement` がある | パッチ（0.3.1 → 0.3.2） | マイナー（1.3.1 → 1.4.0） |
| `bug` がある | パッチ（0.3.1 → 0.3.2） | パッチ（1.3.1 → 1.3.2） |
| `documentation` / `dependencies` だけ | リリースしない | リリースしない |

リリースする時機:

- **リリースも、マージの提案と一緒に版を示し、利用者の了承を得てからタグを push する。**
- **`bug` の PR をマージしたら、すぐにパッチを出す。** 不具合で困っている利用者に早く届けるため。
- **`enhancement` は、利用者に頼まれた一連の作業の PR をすべてマージしてから、まとめて 1 回出す。** PR ごとに版を分けない。
- **`breaking` を含むリリースは、利用者の了承を得てから出す。** 互換を壊す変更はなるべく 1 回のリリースにまとめる。

手順:

```
git fetch origin --tags
$prev = git describe --tags --abbrev=0 origin/main        # 前の版
$since = git log -1 --format=%cI $prev                     # その日時
gh pr list --state merged --base main --search "merged:>$since" --json number,title,labels
```

一覧のラベルで次の版を決め、タグを付けて push する。push したら `gh run watch` で `release` の成功を確かめ、利用者に版と含めた PR の番号を報告する。

```
git tag v0.1.0 origin/main
git push origin v0.1.0
```

- タグは一度 push したら付け直さない。間違えたら、直した版を次の番号で出す。
