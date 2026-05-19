Set-StrictMode -Version Latest

function Find-PAInFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Find,
        [switch]$UseRegex
    )

    if (-not (Test-Path $Path)) {
        return ''
    }
    $content = Get-Content -Path $Path -Raw
    $pattern = if ($UseRegex) { $Find } else { [regex]::Escape($Find) }
    if ($content -match $pattern) {
        return '1'
    }
    return ''
}

function Replace-PAInFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [AllowNull()][string]$Find,
        [AllowNull()][string]$Replace,
        [switch]$UseRegex
    )

    if (-not (Test-Path $Path)) {
        return $false
    }

    $content = Get-Content -Path $Path -Raw
    if ($UseRegex) {
        $updated = $content -replace $Find, $Replace
    } else {
        $updated = $content.Replace($Find, $Replace)
    }

    $item = Get-Item -Path $Path
    if ($item.IsReadOnly) {
        $item.IsReadOnly = $false
    }
    Set-Content -Path $Path -Value $updated -NoNewline
    return ($updated -ne $content)
}

function Get-PASubstringBetween {
    param(
        [AllowNull()][string]$Input,
        [Parameter(Mandatory)][string]$Start,
        [Parameter(Mandatory)][string]$End
    )

    if ($null -eq $Input) {
        return ''
    }
    $startIdx = $Input.IndexOf($Start)
    if ($startIdx -lt 0) {
        return ''
    }
    $startIdx += $Start.Length
    $endIdx = $Input.IndexOf($End, $startIdx)
    if ($endIdx -lt 0) {
        return $Input.Substring($startIdx)
    }
    return $Input.Substring($startIdx, $endIdx - $startIdx)
}

function Get-PASubstringAfter {
    param(
        [AllowNull()][string]$Input,
        [Parameter(Mandatory)][string]$Start
    )

    if ($null -eq $Input) {
        return ''
    }
    $idx = $Input.IndexOf($Start)
    if ($idx -lt 0) {
        return ''
    }
    return $Input.Substring($idx + $Start.Length)
}

function Get-PAXmlValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$XPath
    )

    if (-not (Test-Path $Path)) {
        return ''
    }

    [xml]$xml = Get-Content -Path $Path -Raw
    $node = $xml.SelectSingleNode($XPath)
    if ($node) {
        return $node.InnerText
    }
    return ''
}

function Set-PAXmlValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$XPath,
        [AllowNull()][string]$Value
    )

    if (-not (Test-Path $Path)) {
        return $false
    }

    [xml]$xml = Get-Content -Path $Path -Raw
    $node = $xml.SelectSingleNode($XPath)
    if (-not $node) {
        return $false
    }

    $node.InnerText = $Value
    $item = Get-Item -Path $Path
    if ($item.IsReadOnly) {
        $item.IsReadOnly = $false
    }
    $xml.Save($Path)
    return $true
}

function Get-PANormalizedSourceControlPath {
    param([AllowNull()][string]$Path)

    return ([string]$Path).Replace('\', '/')
}

function Get-PATrunkPath {
    param([AllowNull()][string]$SourceControlPath)

    $normalized = Get-PANormalizedSourceControlPath -Path $SourceControlPath
    return [regex]::Replace($normalized, '/delphi$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function Get-PAVisualStudioPath {
    param([AllowNull()][string]$SourceControlPath)

    $trunk = Get-PATrunkPath -SourceControlPath $SourceControlPath
    return "$trunk/Visual Studio"
}

function Convert-PACarriageReturnsToCrLf {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) {
        return ''
    }
    return $Text -replace "`r(?!`n)", "`r`n"
}

Export-ModuleMember -Function Find-PAInFile, Replace-PAInFile, Get-PASubstringBetween, Get-PASubstringAfter, Get-PAXmlValue, Set-PAXmlValue, Get-PANormalizedSourceControlPath, Get-PATrunkPath, Get-PAVisualStudioPath, Convert-PACarriageReturnsToCrLf
