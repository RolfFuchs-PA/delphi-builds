Set-StrictMode -Version Latest

function New-PASignFileCommand {
    param(
        [Parameter(Mandatory)][string]$SignScriptPath,
        [Parameter(Mandatory)][string]$FilePath,
        [string]$Description = '',
        [string]$PowerShellPath = ''
    )

    if ([string]::IsNullOrWhiteSpace($PowerShellPath)) {
        $pwshPath = 'C:\Program Files\PowerShell\7\pwsh.exe'
        if (Test-Path $pwshPath) {
            $PowerShellPath = $pwshPath
        } else {
            $PowerShellPath = 'powershell.exe'
        }
    }

    return "`"$PowerShellPath`" -NoProfile -ExecutionPolicy Bypass -File `"$SignScriptPath`" -FilePath `"$FilePath`" -Description `"$Description`""
}

function New-PASignVerifyCommand {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [string]$SignToolPath = 'C:\Compilers\SignTool\signtool.exe'
    )

    return "`"$SignToolPath`" verify /pa `"$FilePath`""
}

function Invoke-PASignFile {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [string]$Description = '',
        [string]$SignScriptPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'signfile.ps1'),
        [scriptblock]$Executor = { param($Command) cmd.exe /c $Command }
    )

    $signCommand = New-PASignFileCommand -SignScriptPath $SignScriptPath -FilePath $FilePath -Description $Description
    & $Executor $signCommand
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to sign file: $FilePath"
    }

    $verifyCommand = New-PASignVerifyCommand -FilePath $FilePath
    & $Executor $verifyCommand
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to verify signature: $FilePath"
    }
}

Export-ModuleMember -Function New-PASignFileCommand, New-PASignVerifyCommand, Invoke-PASignFile
