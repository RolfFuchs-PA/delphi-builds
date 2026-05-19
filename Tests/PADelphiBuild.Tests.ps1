$modulePath = Join-Path $PSScriptRoot '..\Modules\PADelphiBuild.psm1'
Import-Module $modulePath -Force -DisableNameChecking

Describe 'PADelphiBuild' {
    It 'parses dcc config defines and unit search paths' {
        $path = Join-Path $TestDrive 'dcc32.cfg'
        @'
-DDEBUG;TRACE
-u"C:\Lib;C:\Other"
'@ | Set-Content -Path $path

        $config = Get-PADccConfigEntries -Path $path
        $config.Defines[0] | Should Be 'DEBUG;TRACE'
        $config.UnitSearchPath[0] | Should Be 'C:\Lib;C:\Other'
    }

    It 'creates legacy dcc command strings' {
        $command = New-PALegacyDccCommand -ProjectFile 'C:\Build\Source\App.dpr' -OutputDirectory 'C:\Build' -UnitSearchPath @('C:\Lib') -Defines @('RELEASE')
        $command | Should Match 'dcc32\.exe'
        $command | Should Match '-U"C:\\Lib"'
        $command | Should Match '-D"RELEASE"'
    }

    It 'creates msbuild command strings' {
        $command = New-PAMsBuildCommand -ProjectName 'App.dproj' -Configuration Debug -Platform Win64 -OutputDirectory '..\debug' -LogFile 'build.log' -Properties @{ DCC_DebugInformation = 2 }
        $command | Should Match 'msbuild "App\.dproj"'
        $command | Should Match 'Config=Debug'
        $command | Should Match 'platform=Win64'
        $command | Should Match 'DCC_DebugInformation=2'
        $command | Should Match '> "build\.log"'
    }
}
