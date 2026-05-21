Set-StrictMode -Version Latest

function ConvertTo-PAVersionParts {
    param([AllowNull()][string]$Version)

    $parts = @(0, 0, 0, 0)
    if ([string]::IsNullOrWhiteSpace($Version)) {
        return $parts
    }

    $rawParts = $Version -split '\.'
    for ($i = 0; $i -lt 4 -and $i -lt $rawParts.Count; $i++) {
        $value = 0
        if ([int]::TryParse($rawParts[$i], [ref]$value)) {
            $parts[$i] = $value
        }
    }

    return $parts
}

function Compare-PAVersion {
    param(
        [AllowNull()][string]$Version1,
        [AllowNull()][string]$Version2
    )

    $v1Parts = ConvertTo-PAVersionParts -Version $Version1
    $v2Parts = ConvertTo-PAVersionParts -Version $Version2

    for ($i = 0; $i -lt 4; $i++) {
        if ($v1Parts[$i] -lt $v2Parts[$i]) { return -1 }
        if ($v1Parts[$i] -gt $v2Parts[$i]) { return 1 }
    }

    return 0
}

function Split-PAVersion {
    param([AllowNull()][string]$Version)

    $parts = ConvertTo-PAVersionParts -Version $Version
    return [pscustomobject]@{
        Major   = $parts[0]
        Minor   = $parts[1]
        Release = $parts[2]
        Build   = $parts[3]
    }
}

function Join-PAVersion {
    param(
        [int]$Major,
        [int]$Minor,
        [int]$Release,
        [int]$Build
    )

    return "$Major.$Minor.$Release.$Build"
}

function Invoke-PAVersionIncrement {
    param(
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][ValidateSet('None', 'Build', 'Release', 'Minor', 'Major')][string]$Part,
        [int]$CurrentYear = (Get-Date).Year
    )

    $parsed = Split-PAVersion -Version $Version
    switch ($Part) {
        'None'    { }
        'Build'   { $parsed.Build++ }
        'Release' { $parsed.Release++; $parsed.Build = 0 }
        'Minor'   { $parsed.Minor++; $parsed.Release = 0; $parsed.Build = 0 }
        'Major'   { $parsed.Major = $CurrentYear; $parsed.Minor = 1; $parsed.Release = 0; $parsed.Build = 0 }
    }

    return Join-PAVersion -Major $parsed.Major -Minor $parsed.Minor -Release $parsed.Release -Build $parsed.Build
}

function Get-PADofVersion {
    param([Parameter(Mandatory)][string]$Path)

    $values = @{
        Major   = 0
        Minor   = 0
        Release = 0
        Build   = 0
    }

    if (-not (Test-Path $Path)) {
        throw "DOF file not found: $Path"
    }

    $inVersionSection = $false
    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        if ($line -match '^\[(.+)\]$') {
            $inVersionSection = ($Matches[1] -eq 'Version Info')
            continue
        }

        if (-not $inVersionSection) {
            continue
        }

        if ($line -match '^(MajorVer|MinorVer|Release|Build)\s*=\s*(\d+)') {
            switch ($Matches[1]) {
                'MajorVer' { $values.Major = [int]$Matches[2] }
                'MinorVer' { $values.Minor = [int]$Matches[2] }
                'Release'  { $values.Release = [int]$Matches[2] }
                'Build'    { $values.Build = [int]$Matches[2] }
            }
        }
    }

    return [pscustomobject]$values
}

function Get-PADprojVersion {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path $Path)) {
        throw "DPROJ file not found: $Path"
    }

    [xml]$xml = Get-Content -Path $Path -Raw
    $values = @{
        Major   = 0
        Minor   = 0
        Release = 0
        Build   = 0
    }

    $modernNodeMap = @{
        Major   = "//*[local-name()='VerInfo_MajorVer']"
        Minor   = "//*[local-name()='VerInfo_MinorVer']"
        Release = "//*[local-name()='VerInfo_Release']"
        Build   = "//*[local-name()='VerInfo_Build']"
    }

    $hasModernVersionInfo = $false
    $hasModernFileVersion = $false
    $verInfoKeysNodes = @($xml.SelectNodes("//*[local-name()='VerInfo_Keys']"))
    foreach ($verInfoKeysNode in $verInfoKeysNodes) {
        $hasModernVersionInfo = $true
        if ($verInfoKeysNode.InnerText -match '(?i)(?:^|;)FileVersion=(\d+(?:\.\d+){1,3})(?:;|$)') {
            $candidateVersion = $Matches[1]
            $currentVersion = Join-PAVersion -Major $values.Major -Minor $values.Minor -Release $values.Release -Build $values.Build
            if (-not $hasModernFileVersion -or (Compare-PAVersion -Version1 $candidateVersion -Version2 $currentVersion) -gt 0) {
                $fileVersionParts = ConvertTo-PAVersionParts -Version $candidateVersion
                $values.Major = $fileVersionParts[0]
                $values.Minor = $fileVersionParts[1]
                $values.Release = $fileVersionParts[2]
                $values.Build = $fileVersionParts[3]
                $hasModernFileVersion = $true
            }
        }
    }

    if (-not $hasModernFileVersion) {
        foreach ($key in $modernNodeMap.Keys) {
            $node = $xml.SelectSingleNode($modernNodeMap[$key])
            if ($node) {
                $hasModernVersionInfo = $true
                $value = 0
                if ([int]::TryParse($node.InnerText, [ref]$value)) {
                    $values[$key] = $value
                }
            }
        }
    }

    if (-not $hasModernVersionInfo) {
        $legacyNodeMap = @{
            Major   = "//*[local-name()='VersionInfo' and @Name='MajorVer']"
            Minor   = "//*[local-name()='VersionInfo' and @Name='MinorVer']"
            Release = "//*[local-name()='VersionInfo' and @Name='Release']"
            Build   = "//*[local-name()='VersionInfo' and @Name='Build']"
        }

        foreach ($key in $legacyNodeMap.Keys) {
            $node = $xml.SelectSingleNode($legacyNodeMap[$key])
            if ($node) {
                $value = 0
                if ([int]::TryParse($node.InnerText, [ref]$value)) {
                    $values[$key] = $value
                }
            }
        }
    }

    return [pscustomobject]$values
}

function Get-PAProjectVersion {
    param([Parameter(Mandatory)][string]$Path)

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($extension) {
        '.dof'   { return Get-PADofVersion -Path $Path }
        '.dproj' { return Get-PADprojVersion -Path $Path }
        default  { throw "Unsupported project version file extension: $extension" }
    }
}

Export-ModuleMember -Function ConvertTo-PAVersionParts, Compare-PAVersion, Split-PAVersion, Join-PAVersion, Invoke-PAVersionIncrement, Get-PADofVersion, Get-PADprojVersion, Get-PAProjectVersion
