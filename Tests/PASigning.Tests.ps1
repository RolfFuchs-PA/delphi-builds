$modulePath = Join-Path $PSScriptRoot '..\Modules\PASigning.psm1'
Import-Module $modulePath -Force -DisableNameChecking

Describe 'PASigning' {
    It 'builds signfile command using repository script path' {
        $command = New-PASignFileCommand -SignScriptPath 'C:\Repo\signfile.ps1' -FilePath 'C:\Build\App.exe' -Description 'App' -PowerShellPath 'powershell.exe'
        $command | Should Match '-NoProfile'
        $command | Should Match '^"powershell\.exe"'
        $command | Should Match 'signfile\.ps1'
        $command | Should Match '-FilePath "C:\\Build\\App\.exe"'
        $command | Should Match '-Description "App"'
    }

    It 'builds signtool verification command' {
        New-PASignVerifyCommand -FilePath 'C:\Build\App.exe' | Should Be '"C:\Compilers\SignTool\signtool.exe" verify /pa "C:\Build\App.exe"'
    }
}
