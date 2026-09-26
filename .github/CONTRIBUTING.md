# Contributing

English | [日本語](CONTRIBUTING.ja.md)

Bug reports, feature requests and proposed fixes for tebunko are welcome. This page describes how to propose a change and the rules for doing so.

- For questions about how to use the tool, see [SUPPORT.md](SUPPORT.md).
- Report security issues privately as described in [SECURITY.md](SECURITY.md), not in an issue.
- Everyone who takes part follows the [Code of Conduct](CODE_OF_CONDUCT.md).
- Issues and pull requests may be written in English or Japanese. The design documents (`docs/`) and most comments in the code are in Japanese.

## How changes are made

Changes go in this order: **issue → branch → pull request → main**.

1. **Open an issue.** Use the "不具合" (bug) template for bugs and the "機能の要望" (feature request) template for new or changed features. Write the title in the form described in "Commit and pull request titles" below (the templates start the title with `fix: ` or `feat: `). A small change such as a typo fix can go straight to a pull request without an issue.
2. **Create a working branch from main.**
3. **Make the change and run the tests** (see "Tests" below).
4. **Open a pull request.** Before you open it, bring in the latest main with a merge (a branch that does not include the latest main cannot be merged; do not rebase and force-push a branch that you have already pushed). Fill in the pull request template and link the issue with `Closes #<issue number>`. If the change breaks compatibility with the previous version, describe how to migrate. The maintainer adds the labels.
5. **When CI passes and the review is done, the pull request is squash merged.** The pull request title becomes the title of the commit on main, so write it as one line that says what the change does, in the form described in "Commit and pull request titles".

Put only one change in each pull request. Send unrelated fixes as separate pull requests.

### Commit and pull request titles

The first line of a commit message, the pull request title and the issue title follow [Conventional Commits](https://www.conventionalcommits.org/en/v1.0.0/). Write the type in English. Write the description in Japanese or English.

```
<type>(<scope>)!: <description>

feat: Excel の図形の文字を検索できるようにする
fix(gui): fix the result count that stays at 0
feat!: change the index format
```

- The scope (such as `(gui)`) is optional. If you write one, make it a single word with no spaces.
- Add `!` to a change after which the settings file or the index of the previous version can no longer be used as is. Describe how to migrate in the pull request.
- After the type, write one colon and one space.

| Type | Use for |
|---|---|
| `feat` | Adding or changing a feature |
| `fix` | Fixing a bug |
| `docs` | Changes to documents only |
| `refactor` | Rewriting without changing behavior |
| `perf` | Making it faster |
| `test` | Adding or fixing tests only |
| `style` | Formatting only (spaces, line breaks) |
| `build` | How the release is built, dependency updates |
| `ci` | CI, git hooks, development tools |
| `chore` | Anything that fits none of the above |
| `revert` | Reverting an earlier change |

The form is checked automatically. The `commit-msg` hook checks commits (see "Setting up" below), and CI checks pull request titles; a pull request with a title in the wrong form cannot be merged. If an issue title is in the wrong form, a comment explaining how to fix it is added automatically. Messages that git creates on its own (`Merge ...`, `Revert "..."`, `fixup! ...`) are not checked.

### Signed-off-by

Create commits with `git commit -s` so that they carry `Signed-off-by: <name> <email address>`. This shows that you agree to the [DCO (Developer Certificate of Origin)](https://developercertificate.org/) and that you have the right to submit the change under the license of this repository. Use the same email address as the author of the commit.

The `commit-msg` hook stops commits without it. CI checks the commits of a pull request, and a pull request that contains a commit without it cannot be merged. If you pushed commits without it, add it with `git rebase --signoff origin/main` and push with `git push --force-with-lease`.

## Setting up

You need:

- Windows and Windows PowerShell 5.1 (included with Windows)
- Pester 5.9.0 (used to run the tests). Windows includes Pester 3.4, which cannot run these tests. Install 5.9.0 once with `Install-Module Pester -RequiredVersion 5.9.0 -Scope CurrentUser -Force -SkipPublisherCheck` (`-SkipPublisherCheck` is needed because the publisher differs from the included 3.4)
- Microsoft Excel, Word and PowerPoint (only to actually build indexes; the automated tests do not need them)

After you clone the repository, turn on the pre-commit checks (once).

```powershell
.\tools\install_hooks.ps1
```

From then on, the following checks run on every commit. If a check fails, fix the content instead of skipping the check with `--no-verify`.

- Personal information (real names, email addresses, paths that contain a user name) and the text encoding of scripts (`tools\check_commit.ps1 -Staged`)
- The tests for the files you changed (`tools\run_commit_tests.ps1`; for example, `tests\<path>.Tests.ps1` for `scripts\<path>.ps1`. When it cannot tell which tests are affected, it runs all the fast tests. CI runs all the tests)
- The form of the first line of the commit message (`tools\check_commit_message.ps1`; see "Commit and pull request titles" above)
- That `Signed-off-by` is present (`tools\check_signoff.ps1`; see "Signed-off-by" above)

## Tests

```powershell
.\tests\run.ps1                 # The default tests (Unit, Io, Meta)
.\tests\run.ps1 -Tag Office     # Tests that need Excel, Word and PowerPoint
```

- **When you add a feature or change behavior, write tests in the same pull request.** Test the text shown on screen and the decisions about what is allowed (the decision layer) especially well.
- **Keep the coverage at or above the baseline.** CI (`.\tests\run.ps1 -Ci`) fails when the coverage falls below the value in `tests/coverage.baseline` (90.0). If it falls, add tests to bring it back. Do not change the baseline; only the maintainer does. The screen layer is not measured.
- **Do not write tests only to raise the number.** Test behavior at the unit level. If several tests differ only in input and expected value, put them in one `It` with `-TestCases`.
- When you change the screen, start the tool with `tebunko.bat` and check it.

How the tests are organized and what CI does is described in [docs/design/testing/index.md](../docs/design/testing/index.md) (Japanese).

## Coding rules

The tests in `tests/meta/` and the static analysis in CI (PSScriptAnalyzer; no findings of severity Error) check these rules automatically.

- **Save scripts and XAML as UTF-8 with BOM and CRLF line endings** (required by Windows PowerShell 5.1).
- **`scripts/shared/` does not know about individual tools.** Tools do not load each other either.
- **The decision layer** (`*_view.ps1`, `index_name.ps1`, `search_query.ps1`, `text.ps1`) does not touch the screen (WPF). Put the text shown on screen and the decisions about what is allowed here, and write tests for them.
- Load every file you add from an entry point (`shared/shared.ps1`, `tebunko/lib.ps1`, `gui.ps1`, `indexer.ps1`).
- **Do not use network access, dynamic code execution, registry changes or third-party libraries.** Users rely on this when they review the tool before introducing it ([docs/safety/index.md](../docs/safety/index.md), Japanese).

The structure of the source code is described in [docs/design/index.md](../docs/design/index.md) (Japanese).

## Do not include personal information

This repository is public. **Anything that has ever been in the history is as good as published**, so check before you commit.

- In test data and in examples in documents, use made-up names (`山田`, `佐藤`), email addresses at `example.com`, and made-up organizations (`(株)山田商事`, `C:\共有\営業部`).
- Do not write absolute paths that contain a user name (`C:\Users\<user name>\...`). If you need an example, use `C:\Users\test\...`.
- Office files hide the author's name and similar information in places you cannot see by looking at the content. If you regenerate test data, remove it as described in [tests/testdata/README.md](../tests/testdata/README.md) (Japanese).
- Use your GitHub noreply address as the email address of the commit author.

## Documents

- The design documents are in `docs/`. If you change behavior, update the related design document in the same pull request.
- Draw figures with Mermaid or draw.io (`.drawio.png`). Do not use ASCII art made of box-drawing characters.
- The design documents are also published as a [website](https://hsgwa.github.io/tebunko/) (MkDocs). When you add a design document, also add it to `nav` in `tools/mkdocs/mkdocs.yml`. A pull request that changes the design documents gets a comment with a preview URL of the site automatically. How to check the site locally is described in [docs/design/testing/ci.md](../docs/design/testing/ci.md) (Japanese).

## Releases

The maintainer makes the releases. The changes in each version are listed on [GitHub Releases](https://github.com/hsgwa/tebunko/releases).

## License

Changes you submit in a pull request (including code, documents and test data) are published under the same [MIT License](../LICENSE) as this repository. Submit only what you made yourself or what you have the right to publish under the MIT License. If you want to bring in third-party code, images or documents as is, discuss it in an issue first (the tool assumes that it contains no third-party components; see [sbom.cdx.json](../sbom.cdx.json)).
