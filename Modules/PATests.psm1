Set-StrictMode -Version Latest

function Test-PATestLogFailure {
    param([AllowNull()][string]$LogText)

    if ([string]::IsNullOrEmpty($LogText)) {
        return $false
    }

    return ($LogText -like '*FAILURES!!!*' -or
            $LogText -like '*An error has occurred during program execution*' -or
            $LogText -like '*Fatal:*' -or
            $LogText -like '*Error:*')
}

function Get-PAProcessNameFromPath {
    param([Parameter(Mandatory)][string]$Path)

    return [System.IO.Path]::GetFileNameWithoutExtension($Path)
}

Export-ModuleMember -Function Test-PATestLogFailure, Get-PAProcessNameFromPath
