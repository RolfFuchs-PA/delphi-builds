param (
    [Parameter(Mandatory = $true)][string] $Repository,
    [Parameter(Mandatory = $true)][string] $OutputFolder
)

# switch to 64bit if running in 32bit mode
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Write-Warning "Changing from 32bit to 64bit PowerShell..."
    $powershell=$PSHOME.tolower().replace("syswow64","sysnative").replace("system32","sysnative")

    if ($myInvocation.Line) {
        &"$powershell\powershell.exe" -NonInteractive -NoProfile $myInvocation.Line
    } else {
        &"$powershell\powershell.exe" -NonInteractive -NoProfile -file "$($myInvocation.InvocationName)" $args
    }

    exit $lastexitcode
}

if (!([System.IO.Directory]::Exists($OutputFolder)))
{
    Write-Error "Folder '$OutputFolder' does not exist"
    exit 1
}

Write-Output "getgithubfiles.ps1: Getting '$Repository' into '$OutputFolder'"

# Read GitHub token from environment variable
$GitHubToken = $env:GITHUB_PASQL_PAT

# Prepare repository URL with authentication if token is available
$repoUrl = $Repository
if ($GitHubToken) {
    # Insert token into URL for private repos
    $repoUrl = $Repository -replace "https://github.com/", "https://${GitHubToken}@github.com/"
}

$tempFolderPath = Join-Path $Env:Temp $(New-Guid)
New-Item -Type Directory -Path $tempFolderPath | Out-Null

try {
    git clone --verbose --depth 1 $repoUrl $tempFolderPath
    
    if ($LASTEXITCODE -ne 0) {
        throw "Git clone failed with exit code $LASTEXITCODE"
    }

    Copy-Item -Path "$tempFolderPath/*" -Destination $OutputFolder -Recurse -Force
}
finally {
    if (Test-Path $tempFolderPath) { 
        Remove-Item -Recurse -Force $tempFolderPath 
    }
}