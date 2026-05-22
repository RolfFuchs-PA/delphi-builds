$modulePath = Join-Path $PSScriptRoot '..\Modules\PAFiles.psm1'
Import-Module $modulePath -Force -DisableNameChecking

Describe 'PAFiles' {
    It 'replaces literal and regex text in files' {
        $path = Join-Path $TestDrive 'text.txt'
        'abc 123' | Set-Content -Path $path -NoNewline

        Replace-PAInFile -Path $path -Find 'abc' -Replace 'xyz' | Should Be $true
        Get-Content -Path $path -Raw | Should Be 'xyz 123'

        Replace-PAInFile -Path $path -Find '\d+' -Replace '456' -UseRegex | Should Be $true
        Get-Content -Path $path -Raw | Should Be 'xyz 456'
    }

    It 'clears read-only attributes before replacing file text' {
        $path = Join-Path $TestDrive 'readonly.txt'
        'before' | Set-Content -Path $path -NoNewline
        (Get-Item -Path $path).IsReadOnly = $true

        Replace-PAInFile -Path $path -Find 'before' -Replace 'after' | Should Be $true

        (Get-Item -Path $path).IsReadOnly | Should Be $false
        Get-Content -Path $path -Raw | Should Be 'after'
    }

    It 'finds literal and regex text in files' {
        $path = Join-Path $TestDrive 'find.txt'
        "PA_COPYRIGHT_YEAR = '2026';" | Set-Content -Path $path -NoNewline

        Find-PAInFile -Path $path -Find 'PA_COPYRIGHT_YEAR' | Should Be '1'
        Find-PAInFile -Path $path -Find "PA_COPYRIGHT_YEAR.*= *'2026';" -UseRegex | Should Be '1'
        Find-PAInFile -Path $path -Find "PA_COPYRIGHT_YEAR.*= *'2025';" -UseRegex | Should Be ''
    }

    It 'extracts substrings' {
        Get-PASubstringBetween -Input 'before [value] after' -Start '[' -End ']' | Should Be 'value'
        Get-PASubstringAfter -Input 'prefix:value' -Start ':' | Should Be 'value'
    }

    It 'reads and writes XML node values' {
        $path = Join-Path $TestDrive 'project.xml'
        '<Project><Version>1</Version></Project>' | Set-Content -Path $path

        Get-PAXmlValue -Path $path -XPath '//Version' | Should Be '1'
        Set-PAXmlValue -Path $path -XPath '//Version' -Value '2' | Should Be $true
        Get-PAXmlValue -Path $path -XPath '//Version' | Should Be '2'
    }

    It 'clears read-only attributes before writing XML node values' {
        $path = Join-Path $TestDrive 'readonly-project.xml'
        '<Project><Version>1</Version></Project>' | Set-Content -Path $path
        (Get-Item -Path $path).IsReadOnly = $true

        Set-PAXmlValue -Path $path -XPath '//Version' -Value '2' | Should Be $true

        (Get-Item -Path $path).IsReadOnly | Should Be $false
        Get-PAXmlValue -Path $path -XPath '//Version' | Should Be '2'
    }

    It 'normalizes source control paths' {
        Get-PATrunkPath -SourceControlPath '$\Products\Bank\Trunk\Delphi' | Should Be '$/Products/Bank/Trunk'
        Get-PAVisualStudioPath -SourceControlPath '$\Products\Bank\Trunk\Delphi' | Should Be '$/Products/Bank/Trunk/Visual Studio'
    }

    It 'normalizes bare carriage returns to CRLF' {
        Convert-PACarriageReturnsToCrLf -Text "one`rtwo`r`nthree" | Should Be "one`r`ntwo`r`nthree"
    }
}
