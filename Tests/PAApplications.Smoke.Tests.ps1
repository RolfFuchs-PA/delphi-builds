Describe 'PAApplications module smoke tests' {
    $moduleFiles = Get-ChildItem -Path (Join-Path $PSScriptRoot '..\Modules') -Filter '*.psm1'

    foreach ($moduleFile in $moduleFiles) {
        It "imports $($moduleFile.Name)" {
            { Import-Module $moduleFile.FullName -Force -DisableNameChecking } | Should Not Throw
        }
    }
}
