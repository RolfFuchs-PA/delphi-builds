Set-StrictMode -Version Latest

function Get-PAVaultAuthArgs {
    param(
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][string]$UserName,
        [string]$Password
    )

    $args = @('-host', $HostName, '-user', $UserName)
    if (-not [string]::IsNullOrEmpty($Password)) {
        $args += @('-password', $Password)
    }
    return $args
}

function Split-PAVaultParameters {
    param([AllowNull()][string]$Parameters)

    if ([string]::IsNullOrWhiteSpace($Parameters)) {
        return @()
    }

    return @([regex]::Matches($Parameters, '"([^"]*)"|([^\s]+)') | ForEach-Object {
        if ($_.Groups[1].Success) {
            $_.Groups[1].Value
        } else {
            $_.Groups[2].Value
        }
    })
}

function New-PAVaultCommandArgs {
    param(
        [Parameter(Mandatory)][string]$Command,
        [Parameter(Mandatory)][string]$Repository,
        [Parameter(Mandatory)][string]$HostName,
        [Parameter(Mandatory)][string]$UserName,
        [string]$Password,
        [string[]]$AdditionalArgs = @()
    )

    return @($Command) + (Get-PAVaultAuthArgs -HostName $HostName -UserName $UserName -Password $Password) + @('-repository', $Repository) + $AdditionalArgs
}

function Resolve-PAVaultPathFromWorkingFile {
    param(
        [Parameter(Mandatory)][string]$LocalRoot,
        [Parameter(Mandatory)][string]$RepositoryRoot,
        [Parameter(Mandatory)][string]$FilePath
    )

    $resolvedLocalRoot = [System.IO.Path]::GetFullPath($LocalRoot).TrimEnd('\', '/')
    $resolvedFilePath = [System.IO.Path]::GetFullPath($FilePath)
    if (-not $resolvedFilePath.StartsWith($resolvedLocalRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "File path '$FilePath' is not under local root '$LocalRoot'."
    }

    $relativePath = $resolvedFilePath.Substring($resolvedLocalRoot.Length).TrimStart('\', '/').Replace('\', '/')
    $normalizedRepositoryRoot = $RepositoryRoot.Replace('\', '/').TrimEnd('/')
    return "$normalizedRepositoryRoot/$relativePath"
}

Export-ModuleMember -Function Get-PAVaultAuthArgs, Split-PAVaultParameters, New-PAVaultCommandArgs, Resolve-PAVaultPathFromWorkingFile
