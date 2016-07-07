[string]$localPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
Write-Output $localPath
Set-Location -Path $localPath 
$branch =git branch
if(!("* master" -in $branch))
{
    git checkout master
}
git pull
$status = git status
if("Untracked files:" -in $status)
{
    Write-Output "Found Untracked files: Please run git commit -a -m and git push if want to update repository"
    
    Read-Host
}
Read-Host