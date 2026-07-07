<#
.SYNOPSIS
    Stages all changes, prompts for a commit message, commits, and pushes.
#>

git add .

$staged = git diff --cached --name-only
if (-not $staged) {
    Write-Host "Nothing staged to commit."
    exit 0
}

$message = Read-Host "Commit message"
if (-not $message) {
    Write-Host "Empty commit message, aborting."
    exit 1
}

git commit -m $message
if ($LASTEXITCODE -ne 0) {
    Write-Host "Commit failed, aborting push."
    exit 1
}

git push
