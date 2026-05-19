$modulePath = Join-Path $PSScriptRoot '..\Modules\PATests.psm1'
Import-Module $modulePath -Force -DisableNameChecking

Describe 'PATests' {
    It 'detects failing test logs' {
        Test-PATestLogFailure -LogText 'all good' | Should Be $false
        Test-PATestLogFailure -LogText 'FAILURES!!! one test failed' | Should Be $true
        Test-PATestLogFailure -LogText 'An error has occurred during program execution' | Should Be $true
    }

    It 'gets process name from executable path' {
        Get-PAProcessNameFromPath -Path 'C:\Build\CollectTests.exe' | Should Be 'CollectTests'
    }
}
