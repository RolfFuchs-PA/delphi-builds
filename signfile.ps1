param (
    [Parameter(Mandatory = $true)][string] $FilePath,
    [Parameter(Mandatory = $true)][string] $Description
)

$ErrorActionPreference = 'Stop'

# switch to 64bit if running in 32bit mode
if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    write-warning "changing from 32bit to 64bit PowerShell..."
    $powershell=$PSHOME.tolower().replace("syswow64","sysnative").replace("system32","sysnative")

    if ($myInvocation.Line) {
        &"$powershell\powershell.exe" -NonInteractive -NoProfile $myInvocation.Line
    } else {
        &"$powershell\powershell.exe" -NonInteractive -NoProfile -file "$($myInvocation.InvocationName)" $args
    }

    exit $lastexitcode
}

if (!([System.IO.File]::Exists($FilePath)))
{
    Write-Error "File '$FilePath' does not exist"
    exit 1
}

$windowsPowerShellModules = Join-Path $env:ProgramFiles 'WindowsPowerShell\Modules'
if ((Test-Path $windowsPowerShellModules) -and (($env:PSModulePath -split ';') -notcontains $windowsPowerShellModules)) {
    $env:PSModulePath = "$windowsPowerShellModules;$env:PSModulePath"
}

Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.KeyVault -ErrorAction Stop

$tennant = "e417d5cc-e5d8-4cad-b2cd-c5ef82dea0a0"
$cred = New-Object System.Management.Automation.PSCredential -ArgumentList $env:BUILD_SIGN_P, ($env:BUILD_SIGN_S | ConvertTo-SecureString -AsPlainText -Force)
Connect-AzAccount -ServicePrincipal -Credential $cred -Tenant $tennant

$HSMSigningVaultURL     = Get-AzKeyVaultSecret -VaultName "DevOpsBuildVariables" -Name "HSMSigningVaultURL" -AsPlainText
$HSMSigningClientId     = Get-AzKeyVaultSecret -VaultName "DevOpsBuildVariables" -Name "HSMSigningClientId" -AsPlainText
$HSMSigningClientSecret = Get-AzKeyVaultSecret -VaultName "DevOpsBuildVariables" -Name "HSMSigningClientSecret" -AsPlainText
$HSMSigningCertName     = Get-AzKeyVaultSecret -VaultName "DevOpsBuildVariables" -Name "HSMSigningCertName" -AsPlainText

& AzureSignTool sign `
    -kvt $tennant `
    -kvu $HSMSigningVaultURL `
    -kvi $HSMSigningClientId `
    -kvs $HSMSigningClientSecret `
    -kvc $HSMSigningCertName `
    -tr "http://timestamp.digicert.com" `
    -d $Description `
    -v `
    $FilePath

if ($LASTEXITCODE -ne 0) {
    Write-Error "Failed to sign '$FilePath'. AzureSignTool exited with code $LASTEXITCODE."
    exit $LASTEXITCODE
}

Write-Host "'$FilePath' signed"
exit 0
