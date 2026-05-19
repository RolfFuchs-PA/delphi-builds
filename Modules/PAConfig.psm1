Set-StrictMode -Version Latest

function Expand-PAIniValue {
    param(
        [AllowNull()][string]$Value,
        [hashtable]$Variables = @{}
    )

    if ($null -eq $Value) {
        return ''
    }

    return [regex]::Replace($Value, '%([A-Za-z_][A-Za-z0-9_]*)%', {
        param($match)
        $name = $match.Groups[1].Value
        if ($Variables.ContainsKey($name)) {
            return [string]$Variables[$name]
        }
        $variable = Get-Variable -Name $name -ValueOnly -ErrorAction SilentlyContinue
        if ($null -ne $variable) {
            return [string]$variable
        }
        return $match.Value
    })
}

function Get-PAIniValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Section,
        [Parameter(Mandatory)][string]$Key,
        [hashtable]$Variables = @{}
    )

    if (-not (Test-Path $Path)) {
        return ''
    }

    $inSection = $false
    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        if ($line -match '^\[(.+)\]$') {
            $inSection = ($Matches[1] -eq $Section)
            continue
        }

        if ($inSection -and $line -match "^$([regex]::Escape($Key))\s*=\s*(.*)") {
            return Expand-PAIniValue -Value $Matches[1].Trim() -Variables $Variables
        }
    }

    return ''
}

function Set-PAIniValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Section,
        [Parameter(Mandatory)][string]$Key,
        [AllowNull()][string]$Value
    )

    if (-not (Test-Path $Path)) {
        Set-Content -Path $Path -Value "[$Section]`r`n$Key=$Value" -NoNewline
        return
    }

    $lines = [System.IO.File]::ReadAllLines($Path)
    $result = [System.Collections.Generic.List[string]]::new()
    $inSection = $false
    $sectionFound = $false
    $keyFound = $false

    foreach ($line in $lines) {
        if ($line -match '^\[(.+)\]$') {
            if ($inSection -and -not $keyFound) {
                $result.Add("$Key=$Value")
                $keyFound = $true
            }
            $inSection = ($Matches[1] -eq $Section)
            if ($inSection) {
                $sectionFound = $true
            }
        } elseif ($inSection -and $line -match "^$([regex]::Escape($Key))\s*=") {
            $result.Add("$Key=$Value")
            $keyFound = $true
            continue
        }

        $result.Add($line)
    }

    if (-not $sectionFound) {
        $result.Add("[$Section]")
    }
    if (-not $keyFound) {
        $result.Add("$Key=$Value")
    }

    if (Test-Path $Path) {
        $item = Get-Item -Path $Path
        if ($item.IsReadOnly) {
            $item.IsReadOnly = $false
        }
    }
    [System.IO.File]::WriteAllLines($Path, $result)
}

function Get-PAIniSectionKeys {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Section
    )

    $keys = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path $Path)) {
        return $keys
    }

    $inSection = $false
    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        if ($line -match '^\[(.+)\]$') {
            $inSection = ($Matches[1] -eq $Section)
            continue
        }
        if ($inSection -and $line -match '^(.+?)\s*=') {
            $keys.Add($Matches[1])
        }
    }

    return $keys
}

function Get-PAIniSectionValues {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Section,
        [hashtable]$Variables = @{}
    )

    $values = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path $Path)) {
        return $values
    }

    $inSection = $false
    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        if ($line -match '^\[(.+)\]$') {
            $inSection = ($Matches[1] -eq $Section)
            continue
        }
        if ($inSection -and $line -match '^.+?\s*=\s*(.*)') {
            $values.Add((Expand-PAIniValue -Value $Matches[1].Trim() -Variables $Variables))
        }
    }

    return $values
}

function Get-PADelphiVersion {
    param(
        [string]$IniPath,
        [string]$Section,
        [string]$SourceControlPath,
        [string]$DefaultVersion = '10.4'
    )

    if ($IniPath -and $Section -and (Test-Path $IniPath)) {
        $explicit = Get-PAIniValue -Path $IniPath -Section $Section -Key 'DELPHI_VERSION'
        if (-not [string]::IsNullOrWhiteSpace($explicit)) {
            return $explicit
        }
    }

    $path = [string]$SourceControlPath
    if ($path -match '(?i)delphi\s*6') { return '6' }
    if ($path -match '(?i)delphi\s*2007') { return '2007' }
    if ($path -match '(?i)delphi\s*xe2') { return 'XE2' }
    if ($path -match '(?i)delphi\s*xe6') { return 'XE6' }
    if ($path -match '(?i)(delphi\s*10\.4|sydney)') { return '10.4' }

    return $DefaultVersion
}

Export-ModuleMember -Function Expand-PAIniValue, Get-PAIniValue, Set-PAIniValue, Get-PAIniSectionKeys, Get-PAIniSectionValues, Get-PADelphiVersion
