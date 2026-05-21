$modulePath = Join-Path $PSScriptRoot '..\Modules\PAVersion.psm1'
Import-Module $modulePath -Force -DisableNameChecking

Describe 'PAVersion' {
    It 'compares four-part versions numerically' {
        Compare-PAVersion '23.1.10.0' '23.1.2.99' | Should Be 1
        Compare-PAVersion '23.1.1.0' '23.1.1.0' | Should Be 0
        Compare-PAVersion '22.9.9.9' '23.0.0.0' | Should Be -1
    }

    It 'increments version parts with expected resets' {
        Invoke-PAVersionIncrement -Version '23.1.1.7' -Part Build -CurrentYear 2026 | Should Be '23.1.1.8'
        Invoke-PAVersionIncrement -Version '23.1.1.7' -Part Release -CurrentYear 2026 | Should Be '23.1.2.0'
        Invoke-PAVersionIncrement -Version '23.1.1.7' -Part Minor -CurrentYear 2026 | Should Be '23.2.0.0'
        Invoke-PAVersionIncrement -Version '23.1.1.7' -Part Major -CurrentYear 2026 | Should Be '2026.1.0.0'
    }

    It 'extracts modern DPROJ version nodes' {
        $path = Join-Path $TestDrive 'Modern.dproj'
        @'
<Project>
  <PropertyGroup>
    <VerInfo_MajorVer>23</VerInfo_MajorVer>
    <VerInfo_MinorVer>1</VerInfo_MinorVer>
    <VerInfo_Release>4</VerInfo_Release>
    <VerInfo_Build>9</VerInfo_Build>
  </PropertyGroup>
</Project>
'@ | Set-Content -Path $path

        $version = Get-PAProjectVersion -Path $path
        Join-PAVersion -Major $version.Major -Minor $version.Minor -Release $version.Release -Build $version.Build | Should Be '23.1.4.9'
    }

    It 'does not mix modern DPROJ version keys with stale version values' {
        $path = Join-Path $TestDrive 'MixedModernLegacy.dproj'
        @'
<Project>
  <PropertyGroup>
    <VerInfo_Keys>CompanyName=Professional Advantage Pty. Ltd.;FileVersion=23.1.0.11;ProductVersion=23.1.0;Comments=</VerInfo_Keys>
  </PropertyGroup>
  <PropertyGroup>
    <VerInfo_MajorVer>23</VerInfo_MajorVer>
    <VerInfo_MinorVer>1</VerInfo_MinorVer>
    <VerInfo_Release>1</VerInfo_Release>
    <VerInfo_Keys>CompanyName=Professional Advantage Pty. Ltd.;FileVersion=23.1.1.0;ProductVersion=23.1.1;Comments=</VerInfo_Keys>
  </PropertyGroup>
  <BorlandProject>
    <Delphi.Personality>
      <VersionInfo Name="MajorVer">6</VersionInfo>
      <VersionInfo Name="MinorVer">2</VersionInfo>
      <VersionInfo Name="Release">0</VersionInfo>
      <VersionInfo Name="Build">11</VersionInfo>
    </Delphi.Personality>
  </BorlandProject>
</Project>
'@ | Set-Content -Path $path

        $version = Get-PAProjectVersion -Path $path
        Join-PAVersion -Major $version.Major -Minor $version.Minor -Release $version.Release -Build $version.Build | Should Be '23.1.1.0'
    }

    It 'extracts legacy DOF version values' {
        $path = Join-Path $TestDrive 'Legacy.dof'
        @'
[Version Info]
MajorVer=6
MinorVer=0
Release=15
Build=42
'@ | Set-Content -Path $path

        $version = Get-PAProjectVersion -Path $path
        Join-PAVersion -Major $version.Major -Minor $version.Minor -Release $version.Release -Build $version.Build | Should Be '6.0.15.42'
    }
}
