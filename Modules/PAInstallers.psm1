Set-StrictMode -Version Latest

function New-PAWiseInstallCommand {
    param(
        [Parameter(Mandatory)][string]$WiseExePath,
        [Parameter(Mandatory)][string]$SetupScriptPath,
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$CopyrightYear
    )

    return "`"$WiseExePath`" /d_PA_VERSION_=`"$Version`" /d_PA_COPYRIGHT_YEAR_=`"$CopyrightYear`" /c `"$SetupScriptPath`""
}

function Get-PASetupFileName {
    param(
        [Parameter(Mandatory)][string]$ProjectTitle,
        [Parameter(Mandatory)][string]$Version,
        [switch]$Debug
    )

    $suffix = ''
    if ($Debug) {
        $suffix = ' - debug'
    }
    return "$ProjectTitle Setup v$Version$suffix.exe"
}

Export-ModuleMember -Function New-PAWiseInstallCommand, Get-PASetupFileName
