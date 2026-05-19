$modulePath = Join-Path $PSScriptRoot '..\Modules\PAConfig.psm1'
Import-Module $modulePath -Force -DisableNameChecking

Describe 'PAConfig' {
    It 'reads INI values and expands percent variables' {
        $path = Join-Path $TestDrive 'PAApplications.ini'
        @'
[Collect]
PROJECT_TITLE=Collect
SOURCE_CONTROL_ROOT_PATH=$/Products/
SOURCE_CONTROL_SOURCE_PATH=%SOURCE_CONTROL_ROOT_PATH%%PROJECT_TITLE%/Trunk
'@ | Set-Content -Path $path

        $variables = @{
            SOURCE_CONTROL_ROOT_PATH = '$/Products/'
            PROJECT_TITLE = 'Collect'
        }
        Get-PAIniValue -Path $path -Section 'Collect' -Key 'SOURCE_CONTROL_SOURCE_PATH' -Variables $variables | Should Be '$/Products/Collect/Trunk'
    }

    It 'sets a value in an existing section' {
        $path = Join-Path $TestDrive 'settings.ini'
        "[Build]`r`nVersion=1" | Set-Content -Path $path

        Set-PAIniValue -Path $path -Section 'Build' -Key 'Version' -Value '2'
        Get-PAIniValue -Path $path -Section 'Build' -Key 'Version' | Should Be '2'
    }

    It 'detects Delphi version from explicit INI value before path fallback' {
        $path = Join-Path $TestDrive 'settings.ini'
        "[Collect]`r`nDELPHI_VERSION=XE6" | Set-Content -Path $path

        Get-PADelphiVersion -IniPath $path -Section 'Collect' -SourceControlPath '$/Products/Collect/Delphi 10.4' | Should Be 'XE6'
    }

    It 'detects Delphi version from source control path' {
        Get-PADelphiVersion -SourceControlPath '$/Products/Collect/Delphi XE2/Trunk' | Should Be 'XE2'
        Get-PADelphiVersion -SourceControlPath '$/Products/Collect/Sydney/Trunk' | Should Be '10.4'
    }
}
