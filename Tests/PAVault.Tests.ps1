$modulePath = Join-Path $PSScriptRoot '..\Modules\PAVault.psm1'
Import-Module $modulePath -Force -DisableNameChecking

Describe 'PAVault' {
    It 'builds auth args with password when supplied' {
        $args = Get-PAVaultAuthArgs -HostName 'sdg1.pa.com.au' -UserName 'autobuild' -Password 'secret'
        $args -join ' ' | Should Be '-host sdg1.pa.com.au -user autobuild -password secret'
    }

    It 'splits quoted custom parameters' {
        $args = Split-PAVaultParameters '-destpath "C:\Build Temp" "$/Products/App"'
        $args.Count | Should Be 3
        $args[1] | Should Be 'C:\Build Temp'
    }

    It 'builds command args' {
        $args = New-PAVaultCommandArgs -Command get -Repository SDG -HostName host -UserName user -Password pass -AdditionalArgs @('$/Products/App')
        $args[0] | Should Be 'get'
        $args | Should Contain '-repository'
        $args | Should Contain '$/Products/App'
    }
}
