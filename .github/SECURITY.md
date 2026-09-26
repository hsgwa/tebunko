# Security

English | [日本語](SECURITY.ja.md)

This page collects information about the safety of tebunko and explains how to contact us when you find a problem.

## How the tool stays safe

What the tool does and does not do, why it does not use dangerous operations or libraries, and the results of checks by third-party tools are described in [docs/safety/index.md](../docs/safety/index.md) (Japanese). If you are reviewing the tool before introducing it, read this document first.

The claims can be checked automatically with the following command (the checks are in `tests/meta/safety.Tests.ps1`).

```powershell
.\tests\run.ps1 -Tag Meta
```

## Reporting a vulnerability or suspicious behavior

If you find a security problem (unintended changes to or deletion of files, unexpected network access, leaks of information and so on), **do not write it in a public issue**. Contact us through GitHub's private vulnerability reporting instead.

**[Report privately](https://github.com/hsgwa/tebunko/security/advisories/new)** (the Security tab of the repository → [Report a vulnerability])

Only the maintainer and the reporter can see the report. When the fix is published, the reporter's name is credited or not as the reporter wishes. Reports in English or Japanese are welcome.

Please include:

- What happened (which file, which operation, and what went wrong)
- Steps to reproduce (conditions of the crawled folder, types of files and so on)
- Your environment (Windows version, Windows PowerShell version, Office version)
- Related logs (`work\インデックス作成ログ.txt` and the error message shown on the screen; hide any confidential information in them)

**Do not send the Office files themselves, because they may contain confidential information.** Attach a file only if you can make a minimal file that causes the same problem.

## Supported versions

Security fixes are made for the latest minor version (the latest `v<major>.<minor>.x`) and published as a patch version. Older versions do not get fixes. Use the latest version from [Releases](https://github.com/hsgwa/tebunko/releases/latest).

| Version | Security fixes |
|---|---|
| Latest minor version | Yes |
| Older versions | No |

## How reports are handled

- We acknowledge a report within 7 days.
- We check the report and try to reproduce the problem.
- If we fix it, we describe the cause and the impact in [docs/safety/index.md](../docs/safety/index.md) (Japanese) and the related design documents, and add tests to `tests/` to prevent it from happening again.
- The fix is released with the catalog and hash list made by `tools\new_release_files.ps1`, and the provenance of the zip is signed with Sigstore. Anyone who receives it can check that the release has not been altered (see the "安全性" (safety) section of the [README](../README.md#安全性), Japanese).

## What we ask of users

The tool **keeps the text of the original documents in plain text in the index (the TSV files in `work\index\`)**. The access rights of the original files are not carried over. Please note the following (details in [docs/safety/disclosure.md](../docs/safety/disclosure.md), Japanese).

- Restrict access to the folder where the tool is placed to at least the same level as the crawled folders.
- Do not put the index as is in a shared folder that users with different access rights can read.
- If you copy the index to another PC, treat it as taking the data out, and follow your organization's rules.
