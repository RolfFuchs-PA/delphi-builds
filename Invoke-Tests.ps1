param(
    [string]$Path = (Join-Path $PSScriptRoot 'Tests')
)

$ErrorActionPreference = 'Stop'

$pester = Get-Module -ListAvailable -Name Pester | Sort-Object Version -Descending | Select-Object -First 1
if (-not $pester) {
    throw 'Pester is not installed. Install or import Pester to run the test suite.'
}

Import-Module Pester -Force
Invoke-Pester -Path $Path
