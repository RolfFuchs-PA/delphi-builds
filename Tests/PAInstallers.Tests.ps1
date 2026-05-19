$modulePath = Join-Path $PSScriptRoot '..\Modules\PAInstallers.psm1'
Import-Module $modulePath -Force -DisableNameChecking

Describe 'PAInstallers' {
    It 'builds Wise command line' {
        $command = New-PAWiseInstallCommand -WiseExePath 'C:\Wise\Wise32.exe' -SetupScriptPath 'C:\Build\Setup\Standard Setup.wse' -Version '23.1.1.8' -CopyrightYear '2026'
        $command | Should Match '"C:\\Wise\\Wise32\.exe"'
        $command | Should Match '/d_PA_VERSION_="23\.1\.1\.8"'
        $command | Should Match '/d_PA_COPYRIGHT_YEAR_="2026"'
        $command | Should Match '/c "C:\\Build\\Setup\\Standard Setup\.wse"'
    }

    It 'builds release and debug setup names' {
        Get-PASetupFileName -ProjectTitle 'Collect' -Version '23.1.1.8' | Should Be 'Collect Setup v23.1.1.8.exe'
        Get-PASetupFileName -ProjectTitle 'Collect' -Version '23.1.1.8' -Debug | Should Be 'Collect Setup v23.1.1.8 - debug.exe'
    }
}
