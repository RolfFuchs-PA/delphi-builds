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

Write-Output "getdevopsfiles.ps1: Getting '$Repository' into '$OutputFolder'"

# foreach ($envVar in Get-ChildItem Env:) {
#     Write-Output "$($envVar.Name) = $($envVar.Value)"
# }

$tennant = "e417d5cc-e5d8-4cad-b2cd-c5ef82dea0a0"
$cred = New-Object System.Management.Automation.PSCredential -ArgumentList $env:BUILD_SIGN_P, ($env:BUILD_SIGN_S | ConvertTo-SecureString -AsPlainText -Force)
Connect-AzAccount -ServicePrincipal -Credential $cred -Tenant $tennant

$tempFolderPath = Join-Path $Env:Temp $(New-Guid); New-Item -Type Directory -Path $tempFolderPath | Out-Null
if (Test-Path $tempFolderPath) { Remove-Item -Recurse -Force $tempFolderPath }
New-Item -Type Directory -Path $tempFolderPath | Out-Null

git clone --verbose --depth 1 $Repository $tempFolderPath

Copy-Item -Path "$tempFolderPath/*" -Destination $OutputFolder -Recurse -Force

if (Test-Path $tempFolderPath) { Remove-Item -Recurse -Force $tempFolderPath }
