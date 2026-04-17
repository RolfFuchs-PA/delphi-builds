param (
    [Parameter(Mandatory = $true)][string] $FilePath,
    [Parameter(Mandatory = $true)][string] $Description
)

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

$tennant = "e417d5cc-e5d8-4cad-b2cd-c5ef82dea0a0"
$cred = New-Object System.Management.Automation.PSCredential -ArgumentList $env:BUILD_SIGN_P, ($env:BUILD_SIGN_S | ConvertTo-SecureString -AsPlainText -Force)
Connect-AzAccount -ServicePrincipal -Credential $cred -Tenant $tennant

$HSMSigningVaultURL     = Get-AzKeyVaultSecret -VaultName "DevOpsBuildVariables" -Name "HSMSigningVaultURL" -AsPlainText
$HSMSigningClientId     = Get-AzKeyVaultSecret -VaultName "DevOpsBuildVariables" -Name "HSMSigningClientId" -AsPlainText
$HSMSigningClientSecret = Get-AzKeyVaultSecret -VaultName "DevOpsBuildVariables" -Name "HSMSigningClientSecret" -AsPlainText
$HSMSigningCertName     = Get-AzKeyVaultSecret -VaultName "DevOpsBuildVariables" -Name "HSMSigningCertName" -AsPlainText

if (!(AzureSignTool sign `
        -kvt $tennant `
        -kvu $HSMSigningVaultURL `
        -kvi $HSMSigningClientId `
        -kvs $HSMSigningClientSecret `
        -kvc $HSMSigningCertName `
        -tr "http://timestamp.digicert.com" `
        -d $Description `
        -v `
        $FilePath
)) {
    Write-Error "Failed to sign '$FilePath'"
} else {
    Write-Host "'$FilePath' signed"
}
