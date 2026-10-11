# テストで一時の git リポジトリを作るときの道具。
#
# コミット前のフックの中や、git の環境変数が残ったシェルからテストが動くと、GIT_DIR・GIT_INDEX_FILE などが
# フックを動かしているリポジトリを指し、GIT_AUTHOR_*・GIT_COMMITTER_* がそのコミットの作者を指している。
# そのまま git init やコミットをすると、一時のリポジトリでなく手元のリポジトリを書き換えてしまう。
# 一時のリポジトリを作るテストは、BeforeAll で clearGitEnv、AfterAll で restoreGitEnv を呼び、
# リポジトリは initTestRepo で作る。

# git が参照する環境変数（git rev-parse --local-env-vars）と、作者・コミッターの環境変数の名前
function gitEnvNames {
    # git が名前を返さなかったときも外せるよう、最低限の名前は固定で持つ
    $names = @(git rev-parse --local-env-vars) | Where-Object { $_ }
    $names += @("GIT_DIR", "GIT_WORK_TREE", "GIT_INDEX_FILE", "GIT_COMMON_DIR", "GIT_PREFIX", "GIT_OBJECT_DIRECTORY")
    $names += @("GIT_AUTHOR_NAME", "GIT_AUTHOR_EMAIL", "GIT_AUTHOR_DATE", "GIT_COMMITTER_NAME", "GIT_COMMITTER_EMAIL", "GIT_COMMITTER_DATE")
    return @($names | Select-Object -Unique)
}

# 外した環境変数の値を返す（restoreGitEnv に渡して戻す）
function clearGitEnv {
    $saved = @{}
    foreach ($name in (gitEnvNames)) {
        $saved[$name] = [Environment]::GetEnvironmentVariable($name)
        [Environment]::SetEnvironmentVariable($name, $null)
    }
    return $saved
}

function restoreGitEnv([hashtable]$saved) {
    foreach ($name in $saved.Keys) {
        [Environment]::SetEnvironmentVariable($name, $saved[$name])
    }
}

# $path に一時のリポジトリを作る。$root（TestDrive）の外を指すなら、コミットを作る前に throw する
function initTestRepo([string]$path, [string]$root) {
    foreach ($name in (gitEnvNames)) {
        if ([Environment]::GetEnvironmentVariable($name)) {
            throw "git の環境変数 $name が残っています。clearGitEnv で外してから一時のリポジトリを作ってください。"
        }
    }
    $base = [System.IO.Path]::GetFullPath($root).TrimEnd("\", "/")
    if (![System.IO.Path]::GetFullPath($path).StartsWith($base + "\", [StringComparison]::OrdinalIgnoreCase)) {
        throw "一時のリポジトリの場所 $path が $root の外です。"
    }
    git init -q $path
    $dir = [System.IO.Path]::GetFullPath(((git -C $path rev-parse --absolute-git-dir) -join ""))
    if (!$dir.StartsWith($base + "\", [StringComparison]::OrdinalIgnoreCase)) {
        throw "一時のリポジトリの .git が $root の外にあります。"
    }
}
