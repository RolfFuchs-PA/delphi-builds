Set-StrictMode -Version Latest

function Get-PADccConfigEntries {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path $Path)) {
        throw "DCC config file not found: $Path"
    }

    $defines = [System.Collections.Generic.List[string]]::new()
    $unitSearch = [System.Collections.Generic.List[string]]::new()

    foreach ($line in Get-Content -Path $Path) {
        $current = $line.Trim()
        if ($current.StartsWith('-D')) {
            $defines.Add($current.Substring(2).Trim().Trim('"'))
        } elseif ($current.StartsWith('-u')) {
            $unitSearch.Add($current.Substring(2).Trim().Trim('"'))
        }
    }

    return [pscustomobject]@{
        Defines        = @($defines)
        UnitSearchPath = @($unitSearch)
    }
}

function New-PALegacyDccCommand {
    param(
        [Parameter(Mandatory)][string]$ProjectFile,
        [Parameter(Mandatory)][string]$OutputDirectory,
        [string[]]$UnitSearchPath = @(),
        [string[]]$Defines = @(),
        [switch]$Debug
    )

    $args = [System.Collections.Generic.List[string]]::new()
    $args.Add('dcc32.exe')
    $args.Add('-B')
    $args.Add("-E`"$OutputDirectory`"")
    if ($UnitSearchPath.Count -gt 0) {
        $args.Add("-U`"$($UnitSearchPath -join ';')`"")
    }
    if ($Defines.Count -gt 0) {
        $args.Add("-D`"$($Defines -join ';')`"")
    }
    if ($Debug) {
        $args.Add('-V')
    }
    $args.Add("`"$ProjectFile`"")

    return ($args -join ' ')
}

function New-PAMsBuildCommand {
    param(
        [Parameter(Mandatory)][string]$ProjectName,
        [ValidateSet('Release', 'Debug')][string]$Configuration = 'Release',
        [ValidateSet('Win32', 'Win64')][string]$Platform = 'Win32',
        [string]$OutputDirectory = '..\',
        [string]$LogFile,
        [hashtable]$Properties = @{}
    )

    $allProperties = [ordered]@{
        Config         = $Configuration
        platform       = $Platform
        EnvOptionsWarn = 'false'
        DCC_ExeOutput  = $OutputDirectory
    }

    foreach ($key in $Properties.Keys) {
        $allProperties[$key] = $Properties[$key]
    }

    $propertyString = ($allProperties.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join ';'
    $command = "msbuild `"$ProjectName`" /verbosity:diag /clp:ShowCommandLine /t:Rebuild /p:$propertyString"
    if ($LogFile) {
        $command = "$command > `"$LogFile`""
    }
    return $command
}

function Get-PADelphiCompilerRoot {
    param([Parameter(Mandatory)][string]$DelphiVersion)

    switch ($DelphiVersion.ToUpperInvariant()) {
        '6'    { return 'C:\Compilers\Delphi 6' }
        '2007' { return 'C:\Compilers\RAD Studio 2007\CodeGear' }
        'XE2'  { return 'C:\Compilers\Delphi XE2\Embarcadero' }
        'XE6'  { return 'C:\Compilers\Delphi XE6\Embarcadero' }
        '10.4' { return 'C:\Compilers\Delphi 10.4\Embarcadero' }
        default { throw "Unsupported Delphi version: $DelphiVersion" }
    }
}

Export-ModuleMember -Function Get-PADccConfigEntries, New-PALegacyDccCommand, New-PAMsBuildCommand, Get-PADelphiCompilerRoot
