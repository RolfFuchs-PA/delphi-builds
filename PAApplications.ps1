#Requires -Version 7.0
<#
.SYNOPSIS
    PA Applications Build Script
    Converted from SmartBear BuildStudio (PAApplications.bxp)

.DESCRIPTION
    Builds PA Applications Delphi projects, runs tests, builds setups,
    and manages Vault source control operations.

.PARAMETER INI_SECTION
    Project name to build (e.g., 'Bank Reconciliation', 'JET').
    If not provided, interactive selection is shown.

.PARAMETER NIGHTLY_BUILD
    Set to TRUE for nightly/unattended builds.
#>
[CmdletBinding()]
param(
    [string]$INI_SECTION = '',
    [string]$NIGHTLY_BUILD = 'FALSE'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# BuildStudio built-in: directory containing this script (with trailing backslash)
$ABSOPENEDPROJECTDIR = $PSScriptRoot + [IO.Path]::DirectorySeparatorChar

#======================================================================
# CONSTANTS
#======================================================================
$script:BUILD_LOCATION_PREFIX = 'C:\Builds\Under Development'
$script:PA_UNIT_FILE_NAME = '\\%VAULT_SERVER_ADDRESS%\forqa\PA Unit Testing\latest-do-not-delete\PAUnitCMD.exe'
$script:VAULT_SERVER_ADDRESS = 'sdg1.pa.com.au'
$script:COMPILER_XE2_DB_EXPRESS_PATH = '-u"C:\\Compilers\\Delphi XE2\\3rd Party\\dbExpress\\v 7.1\\Lib"'
$script:COMPILER_XE2_DEV_EXPRESS_PATH = '-u"C:\\Compilers\\Delphi XE2\\3rd Party\\Developer Express\\v 20.1.5\\Lib"'
$script:COMPILER_XE2_JEDI_CODE_PATH = '-u"C:\\Compilers\\Delphi XE2\\3rd Party\\JEDI Code Library\\v 2.3\\win32"'
$script:COMPILER_XE2_REPORT_BUILDER_PATH = '-u"C:\\Compilers\\Delphi XE2\\3rd Party\\Report Builder Enterprise\\v 20.03\\Lib\\Win32"'
$script:COMPILER_XE2_CLEVER_INTERNET = '-u"C:\\Compilers\\Delphi XE2\\3rd Party\\Clever Internet Suite\\v 9.6\\Lib"'
$script:COMPILER_XE2_PGP_PATH = '-u"C:\\Compilers\\Delphi XE2\\3rd Party\\SecureBlackbox\\v 12.0\\Units\\Delphi16\\Win32"'
$script:COMPILER_XE6_DB_EXPRESS_PATH = '-u"C:\\Compilers\\Delphi XE6\\3rd Party\\dbExpress\\v 7.1\\Lib\\Win64"'
$script:COMPILER_XE6_DEV_EXPRESS_PATH = '-u"C:\\Compilers\\Delphi XE6\\3rd Party\\Developer Express\\v 22.1.3\\Lib\\Win64"'
$script:COMPILER_XE6_JEDI_CODE_PATH = '-u"C:\\Compilers\\Delphi XE6\\3rd Party\\Jedi Code Library\\v 2.6.0\\lib\\d20\\Win64"'
$script:COMPILER_XE6_REPORT_BUILDER_PATH = '-u"C:\\Compilers\\Delphi XE6\\3rd Party\\Report Builder Enterprise\\v 20.03\\Lib\\Win64"'
$script:COMPILER_XE6_CLEVER_INTERNET = '-u"C:\\Compilers\\Delphi XE6\\3rd Party\\Clever Internet Suite\\v 9.6\\Lib"'
$script:COMPILER_XE6_PGP_PATH = '-u"C:\\Compilers\\Delphi XE6\\3rd Party\\SecureBlackbox\\v 12.0\\Units\\DelphiXE6\\Win64"'
$script:COMPILER_104_DB_EXPRESS_32_PATH = '-u"C:\\Compilers\\Delphi 10.4\\3rd Party\\dbExpress\\v 8.3\\Win32"'
$script:COMPILER_104_DEV_EXPRESS_32_PATH = '-u"C:\\Compilers\\Delphi 10.4\\3rd Party\\Developer Express\\v 24.1.6\\Lib"'
$script:COMPILER_104_JEDI_CODE_32_PATH = '-u"C:\\Compilers\\Delphi 10.4\\3rd Party\\Jedi Code Library\\v 2.8\\Win32"'
$script:COMPILER_104_REPORT_BUILDER_32_PATH = '-u"C:\\Compilers\\Delphi 10.4\\3rd Party\\Report Builder Enterprise\\v 22.01\\Lib\\Win32"'
$script:COMPILER_104_CLEVER_INTERNET_32_PATH = '-u"C:\\Compilers\\Delphi 10.4\\3rd Party\\Clever Internet Suite\\v 9.6\\Lib"'
$script:COMPILER_104_DB_EXPRESS_64_PATH = '-u"C:\\Compilers\\Delphi 10.4\\3rd Party\\dbExpress\\v 8.3\\Win64"'
$script:COMPILER_104_DEV_EXPRESS_64_PATH = '-u"C:\\Compilers\\Delphi 10.4\\3rd Party\\Developer Express\\v 24.1.6\\Lib\\Win64"'
$script:COMPILER_104_JEDI_CODE_64_PATH = '-u"C:\\Compilers\\Delphi 10.4\\3rd Party\\Jedi Code Library\\v 2.8\\Win64"'
$script:COMPILER_104_REPORT_BUILDER_64_PATH = '-u"C:\\Compilers\\Delphi 10.4\\3rd Party\\Report Builder Enterprise\\v 22.01\\Lib\\Win64"'
$script:COMPILER_104_CLEVER_INTERNET_64_PATH = '-u"C:\\Compilers\\Delphi 10.4\\3rd Party\\Clever Internet Suite\\v 9.6\\Lib64"'
$script:PA_VS_FRAMEWORK_PATH = '$/Framework/VisualStudio/PA.FrameWork/Trunk'
$script:PA_VS_BLAZOR_FRAMEWORK_PATH = '$/Framework/VisualStudio/PA.Blazor.Components'
$script:MSBUILD_EXE = '"%ProgramFiles(x86)%\MSBuild\14.0\bin\msbuild.exe"'
$script:NUGET_EXE = '"%ProgramFiles(x86)%\MSBuild\14.0\bin\nuget.exe"'
$script:EUREKALOG_PATH = '"C:\Program Files (x86)\Neos Eureka S.r.l\EurekaLog 7\Bin\ecc32speed.exe"'
$script:COMPILER_XE2_EUREKALOG_32BIT_PATH = '-u"C:\\Program Files (x86)\\Neos Eureka S.r.l\\EurekaLog 7\\Lib\\Win32\\Release\\Studio16"'
$script:COMPILER_XE6_EUREKALOG_32BIT_PATH = '-u"C:\\Program Files (x86)\\Neos Eureka S.r.l\\EurekaLog 7\\Lib\\Win32\\Release\\Studio20"'
$script:COMPILER_XE6_EUREKALOG_64BIT_PATH = '-u"C:\\Program Files (x86)\\Neos Eureka S.r.l\\EurekaLog 7\\Lib\\Win64\\Release\\Studio20"'
$script:COMPILER_104_EUREKALOG_32_PATH = '-u"C:\\Compilers\\Delphi 10.4\\3rd Party\\EurekaLog\\v 7.9.2\\Lib\\Win32\\Release\\Studio27"'
$script:COMPILER_104_EUREKALOG_64_PATH = '-u"C:\\Compilers\\Delphi 10.4\\3rd Party\\EurekaLog\\v 7.9.2\\Lib\\Win64\\Release\\Studio27"'
$script:COMPILER_EUREKALOG_COMMON_PATH = 'C:\Program Files (x86)\Neos Eureka S.r.l\EurekaLog 7\Source\Common'  # Added to the search path directly, so no quotes required
$script:DOTNET = '"C:\Program Files (x86)\dotnet\dotnet.exe"'
$script:MSBUILD_2019_EXE = '"C:\Program Files (x86)\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\msbuild.exe"'

#======================================================================
# VARIABLES
#======================================================================
$CURRENT_YEAR = ''  # For copyright etc
$PROJECT_TITLE = ''  # Used to name setup executables
$SOURCE_CONTROL_LABEL = ''  # Base the build on sources matching this label
$HELP_FILE_LOCATION = ''  # Overrides the standard help location
$SOURCE_CONTROL_VERSION = ''  # $/Delphi XYZ/Projects/<project>/SOURCE_CONTROL_VERSION (may not be present)
$SOURCE_CONTROL_ROOT_PATH = ''  # Root folder in Source Control
$SOURCE_CONTROL_PROJECT_PATH = ''  # Project folder in Source Control
$SOURCE_CONTROL_SOURCE_PATH = ''  # Full path to the project in Source Control
$DELPHI_VERSION = ''  # Version of the Delphi compiler required
$BUILD_PATH = ''  # Full path to the build folder
$BUILD_TEMP_PATH = ''  # Temporary files will be built here
$SUPPORTS_ORACLE = ''  # If this flag is not present, Oracle support is assumed to be TRUE
$SUPPORTS_SUN4 = ''  # If this flag is not present, Sun 4 support is assumed to be TRUE
$SUPPORTS_SUN5 = ''  # If this flag is not present, Sun 5 support is assumed to be TRUE
$SUPPORTS_SUN6 = ''  # If this flag is not present, Sun 6 support is assumed to be TRUE
$V_MAJOR = ''
$V_MINOR = ''
$V_RELEASE = ''
$V_BUILD = ''
$BUILD_VERSION = ''  # Generated during building - actual build version N.N.N.N
$RELEASE_VERSION = ''  # Generated during building - \\%VAULT_SERVER_ADDRESS%\Groups\SDG\forqa\<project>\RELEASE_VERSION
$BUILD_INI = ''  # Used to store some information about build
$LAST_BUILD_DATE_TIME = ''  # Date and time of the last successful build
$LAST_BUILD_VERSION = ''  # Full path of the final destination of the last successful build
$VAR_RESULT = ''  # Hold operations result
$VAR_RESULT_TEXT = ''  # Hold operations result text
$PROJECT_LABEL_PATH = ''  # This is where the final build label will be applied
$TEMP_VAR = ''
$TEMP_VAR_2 = ''
$TEMP_VAR_3 = ''
$ICON_FILE_NAME = ''
$NEXT_PROJECT_TO_BUILD = '0'  # In case we need to build multiple projects Advanced/Archive Inquiry

#----------------------------------------------------------------------
# Vault Source Control Credentials
#----------------------------------------------------------------------
$VAULT_USERNAME = $env:VAULT_USERNAME ?? 'autobuild'
$VAULT_PASSWORD = $env:VAULT_PASSWORD ?? 'autobuild'

#======================================================================
# HELPER FUNCTIONS
#======================================================================

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$timestamp] $Message"
    if ($script:LogFile) {
        "[$timestamp] $Message" >> $script:LogFile
    }
}

function Get-IniValue {
    param([string]$Path, [string]$Section, [string]$Key)
    if (-not (Test-Path $Path)) { return '' }
    $inSection = $false
    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        if ($line -match '^\[(.+)\]$') {
            $inSection = ($Matches[1] -eq $Section)
        } elseif ($inSection -and $line -match "^$([regex]::Escape($Key))\s*=\s*(.*)") {
            $raw = $Matches[1].Trim()
            # Expand %VAR% references using current script variables
            $expanded = [regex]::Replace($raw, '%([A-Za-z_][A-Za-z0-9_]*)%', {
                param($m)
                $v = Get-Variable -Name $m.Groups[1].Value -ValueOnly -ErrorAction SilentlyContinue
                if ($null -ne $v) { $v } else { $m.Value }
            })
            return $expanded
        }
    }
    return ''
}

function Set-IniValue {
    param([string]$Path, [string]$Section, [string]$Key, [string]$Value)
    if (-not (Test-Path $Path)) {
        Set-Content -Path $Path -Value "[$Section]`r`n$Key=$Value"
        return
    }
    $lines = [System.IO.File]::ReadAllLines($Path)
    $result = [System.Collections.Generic.List[string]]::new($lines.Length + 2)
    $inSection = $false
    $keyFound = $false
    foreach ($line in $lines) {
        if ($line -match '^\[(.+)\]$') {
            if ($inSection -and -not $keyFound) {
                $result.Add("$Key=$Value")
                $keyFound = $true
            }
            $inSection = ($Matches[1] -eq $Section)
        } elseif ($inSection -and $line -match "^$([regex]::Escape($Key))\s*=") {
            $result.Add("$Key=$Value")
            $keyFound = $true
            continue
        }
        $result.Add($line)
    }
    if (-not $keyFound) {
        if (-not $inSection) { $result.Add("[$Section]") }
        $result.Add("$Key=$Value")
    }
    [System.IO.File]::WriteAllLines($Path, $result)
}

function Get-IniSectionKeys {
    param([string]$Path, [string]$Section)
    $keys = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path $Path)) { return $keys }
    $inSection = $false
    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        if ($line -match '^\[(.+)\]$') {
            $inSection = ($Matches[1] -eq $Section)
        } elseif ($inSection -and $line -match '^(.+?)\s*=') {
            $keys.Add($Matches[1])
        }
    }
    return $keys
}

function Get-IniSectionValues {
    param([string]$Path, [string]$Section)
    $values = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path $Path)) { return $values }
    $inSection = $false
    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        if ($line -match '^\[(.+)\]$') {
            $inSection = ($Matches[1] -eq $Section)
        } elseif ($inSection -and $line -match '^.+?\s*=\s*(.*)') {
            $values.Add($Matches[1].Trim())
        }
    }
    return $values
}

function Find-InFile {
    param([string]$Path, [string]$Find)
    if (-not (Test-Path $Path)) { return '' }
    $content = Get-Content -Path $Path -Raw
    if ($content -match [regex]::Escape($Find)) { return $Find }
    return ''
}

function Replace-InFile {
    param([string]$Path, [string]$Find, [string]$Replace)
    if (-not (Test-Path $Path)) { return }
    $content = Get-Content -Path $Path -Raw
    $content = $content.Replace($Find, $Replace)
    Set-Content -Path $Path -Value $content -NoNewline
}

function Get-SubstringBetween {
    param([string]$Input, [string]$Start, [string]$End)
    $startIdx = $Input.IndexOf($Start)
    if ($startIdx -lt 0) { return '' }
    $startIdx += $Start.Length
    $endIdx = $Input.IndexOf($End, $startIdx)
    if ($endIdx -lt 0) { return $Input.Substring($startIdx) }
    return $Input.Substring($startIdx, $endIdx - $startIdx)
}

function Get-SubstringAfter {
    param([string]$Input, [string]$Start)
    $idx = $Input.IndexOf($Start)
    if ($idx -lt 0) { return '' }
    return $Input.Substring($idx + $Start.Length)
}

function Remove-ItemSafe {
    param([string]$Path, [switch]$Recurse)
    if (Test-Path $Path) {
        if ($Recurse) { Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue }
        else { Remove-Item -Path $Path -Force -ErrorAction SilentlyContinue }
    }
}

function Copy-FileEx {
    param([string]$Source, [string]$Destination, [switch]$Force, [switch]$Recurse)
    $srcDir = Split-Path -Path $Source -Parent
    $srcFilter = Split-Path -Path $Source -Leaf
    if (-not (Test-Path $srcDir)) {
        Write-Log "[WARNING] Source directory not found: $srcDir"
        return $false
    }
    if ($Recurse) {
        Copy-Item -Path $Source -Destination $Destination -Force:$Force -Recurse -ErrorAction Stop
    } else {
        Copy-Item -Path $Source -Destination $Destination -Force:$Force -ErrorAction Stop
    }
    return $true
}

function Invoke-DosCommand {
    param([string]$Command, [string]$WorkingDirectory)
    Write-Log "Executing: $Command"
    $origDir = $PWD
    try {
        if ($WorkingDirectory -and (Test-Path $WorkingDirectory)) {
            Set-Location $WorkingDirectory
        }
        $output = cmd.exe /c $Command 2>&1
        $script:LastExitCode = $LASTEXITCODE
        return ($output -join "`r`n")
    } finally {
        Set-Location $origDir
    }
}

function Invoke-Program {
    param([string]$Path, [string]$Arguments, [string]$WorkingDirectory, [int]$TimeoutSeconds = 0)
    Write-Log "Running: $Path $Arguments"
    $psi = [System.Diagnostics.ProcessStartInfo]@{
        FileName               = $Path
        Arguments              = $Arguments ?? ''
        WorkingDirectory       = $WorkingDirectory ?? $PWD.Path
        UseShellExecute        = $false
        RedirectStandardOutput = $true
        RedirectStandardError  = $true
    }
    $process = [System.Diagnostics.Process]::Start($psi)
    $stdOut = $process.StandardOutput.ReadToEndAsync()
    $stdErr = $process.StandardError.ReadToEndAsync()
    $exited = if ($TimeoutSeconds -gt 0) {
        $process.WaitForExit($TimeoutSeconds * 1000)
    } else {
        $process.WaitForExit(); $true
    }
    if (-not $exited) {
        $process.Kill($true)
        throw "Process timed out after ${TimeoutSeconds}s: $Path"
    }
    [System.Threading.Tasks.Task]::WaitAll($stdOut, $stdErr)
    $script:LAST_STDOUT = $stdOut.Result
    $script:LAST_STDERR = $stdErr.Result
    return $process.ExitCode
}

function Compare-Versions {
    param([string]$Version1, [string]$Version2)
    # Returns: 1 if Version1 > Version2, -1 if Version1 < Version2, 0 if equal
    $v1Parts = $Version1 -split '\.' | ForEach-Object { [int]$_ }
    $v2Parts = $Version2 -split '\.' | ForEach-Object { [int]$_ }
    
    for ($i = 0; $i -lt 4; $i++) {
        $v1 = $v1Parts[$i] ?? 0
        $v2 = $v2Parts[$i] ?? 0
        if ($v1 -lt $v2) { return -1 }
        if ($v1 -gt $v2) { return 1 }
    }
    return 0
}

#region Vault Source Control Helpers

$script:VaultExe = 'C:\Program Files (x86)\SourceGear\Vault Client\vault.exe'
if (-not (Test-Path $script:VaultExe)) {
    $found = Get-Command vault.exe -ErrorAction SilentlyContinue
    $script:VaultExe = $found ? $found.Source : 'vault.exe'
}

function Get-VaultAuthArgs {
    $authArgs = @('-host', $VAULT_SERVER_ADDRESS, '-user', $VAULT_USERNAME)
    if ($VAULT_PASSWORD) { $authArgs += @('-password', $VAULT_PASSWORD) }
    return $authArgs
}

function Invoke-VaultCheckOut {
    param([string]$Repository, [string]$Path, [string]$Host)
    Write-Log "Vault CheckOut: $Path"
    $auth = Get-VaultAuthArgs
    if ($Host) { $auth[1] = $Host }
    & $script:VaultExe checkout @auth -repository $Repository `"$Path`"
    if ($LASTEXITCODE -ne 0) { throw "Vault checkout failed: $Path" }
}

function Invoke-VaultGetLatest {
    param([string]$Repository, [string]$Path, [string]$LocalFolder)
    Write-Log "Vault GetLatest: $Path"
    $auth = Get-VaultAuthArgs
    $destArg = $LocalFolder ? @('-destpath', $LocalFolder) : @()
    & $script:VaultExe get @auth -repository $Repository @destArg `"$Path`"
    if ($LASTEXITCODE -ne 0) { throw "Vault get latest failed: $Path" }
}

function Invoke-VaultCheckIn {
    param([string]$Repository, [string]$Path, [string]$Comment)
    Write-Log "Vault CheckIn: $Path"
    $auth = Get-VaultAuthArgs
    & $script:VaultExe checkin @auth -repository $Repository -comment `"$Comment`" `"$Path`"
    if ($LASTEXITCODE -ne 0) { throw "Vault checkin failed: $Path" }
}

function Invoke-VaultLabel {
    param([string]$Repository, [string]$Path, [string]$Label)
    Write-Log "Vault Label: $Path -> $Label"
    $auth = Get-VaultAuthArgs
    & $script:VaultExe label @auth -repository $Repository `"$Path`" `"$Label`"
}

function Invoke-VaultUndoCheckOut {
    param([string]$Repository, [string]$Path)
    Write-Log "Vault UndoCheckOut: $Path"
    $auth = Get-VaultAuthArgs
    & $script:VaultExe undocheckout @auth -repository $Repository `"$Path`" 2>$null
}

function Invoke-VaultCommand {
    param([string]$Repository, [string]$Command, [string]$Parameters)
    Write-Log "Vault Custom: $Command $Parameters"
    $auth = Get-VaultAuthArgs
    # Split parameters string respecting quoted values
    $paramArgs = @()
    if ($Parameters) {
        $paramArgs = [regex]::Matches($Parameters, '"([^"]*)"|([^\s]+)') | ForEach-Object {
            if ($_.Groups[1].Success) { $_.Groups[1].Value } else { $_.Groups[2].Value }
        }
    }
    & $script:VaultExe $Command @auth -repository $Repository @paramArgs
}

function Invoke-VaultGetByLabel {
    param([string]$Repository, [string]$Path, [string]$Label, [string]$LocalPath)
    Write-Log "Vault GetByLabel: $Path ($Label)"
    $auth = Get-VaultAuthArgs
    $destArg = $LocalPath ? @('-destpath', $LocalPath) : @()
    & $script:VaultExe getlabel @auth -repository $Repository @destArg `"$Path`" `"$Label`"
}

function Get-VaultFiles {
    param([string]$Repository, [string]$Path, [string]$Filter, [int]$TimeoutSeconds = 15)
    $auth = Get-VaultAuthArgs
    $argList = @('listfolder') + $auth + @('-repository', $Repository, $Path)
    $proc = Start-Process -FilePath $script:VaultExe -ArgumentList $argList -NoNewWindow -PassThru -RedirectStandardOutput "$env:TEMP\vault_list.txt" -RedirectStandardError "$env:TEMP\vault_list_err.txt"
    if (-not $proc.WaitForExit($TimeoutSeconds * 1000)) {
        $proc | Stop-Process -Force
        Write-Log "Warning: Vault listfolder timed out after ${TimeoutSeconds}s"
        return @()
    }
    $output = Get-Content "$env:TEMP\vault_list.txt" -ErrorAction SilentlyContinue
    return $Filter ? ($output | Where-Object { $_ -like $Filter }) : $output
}

#endregion Vault Source Control Helpers

function Get-XmlValue {
    param([string]$Path, [string]$XPath)
    if (-not (Test-Path $Path)) { return '' }
    [xml]$xml = Get-Content -Path $Path -Raw
    $node = $xml.SelectSingleNode($XPath)
    if ($node) { return $node.InnerText }
    return ''
}

function Set-XmlValue {
    param([string]$Path, [string]$XPath, [string]$Value)
    if (-not (Test-Path $Path)) { return }
    [xml]$xml = Get-Content -Path $Path -Raw
    $node = $xml.SelectSingleNode($XPath)
    if ($node) {
        $node.InnerText = $Value
        $xml.Save($Path)
    }
}

function Confirm-Action {
    param([string]$Message, [string]$Default = 'Yes')
    Add-Type -AssemblyName System.Windows.Forms
    $result = [System.Windows.Forms.MessageBox]::Show($Message, 'Build Confirmation', 'YesNo', 'Question')
    return ($result -eq 'Yes') ? 'Yes' : 'No'
}

function Show-RadioMenu {
    param([string]$Title, [string[]]$Options, [int]$DefaultIndex = 0)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $font = [System.Drawing.Font]::new('Segoe UI', 9.5)
    $columns = 2
    $radioHeight = 26
    $colWidth = 280
    $padding = 12
    $rows = [Math]::Ceiling($Options.Count / $columns)
    $panelHeight = $rows * $radioHeight + $padding
    $formWidth = $colWidth * $columns + $padding * 3

    $form = [System.Windows.Forms.Form]@{
        Text            = $Title
        StartPosition   = 'CenterScreen'
        FormBorderStyle = 'FixedDialog'
        MaximizeBox     = $false
        MinimizeBox     = $false
        TopMost         = $true
        ClientSize      = [System.Drawing.Size]::new($formWidth, $panelHeight + 50)
        Font            = $font
    }

    $panel = [System.Windows.Forms.Panel]@{
        Location    = [System.Drawing.Point]::new($padding, $padding)
        Size        = [System.Drawing.Size]::new($formWidth - $padding * 2, $panelHeight)
        AutoScroll  = $true
    }

    $radios = [System.Collections.Generic.List[System.Windows.Forms.RadioButton]]::new()
    for ($i = 0; $i -lt $Options.Count; $i++) {
        $col = $i % $columns
        $row = [Math]::Floor($i / $columns)
        $rb = [System.Windows.Forms.RadioButton]@{
            Text     = $Options[$i]
            Location = [System.Drawing.Point]::new($col * $colWidth + 4, $row * $radioHeight)
            Size     = [System.Drawing.Size]::new($colWidth - 8, $radioHeight)
            Checked  = ($i -eq $DefaultIndex)
            Tag      = $i
        }
        $rb.Add_DoubleClick({ $form.DialogResult = 'OK'; $form.Close() })
        $panel.Controls.Add($rb)
        $radios.Add($rb)
    }

    $btnPanel = [System.Windows.Forms.Panel]@{
        Dock   = 'Bottom'
        Height = 40
    }
    $okButton = [System.Windows.Forms.Button]@{
        Text         = 'OK'
        DialogResult = 'OK'
        Size         = [System.Drawing.Size]::new(80, 28)
    }
    $okButton.Location = [System.Drawing.Point]::new($formWidth - 180, 6)
    $cancelButton = [System.Windows.Forms.Button]@{
        Text         = 'Cancel'
        DialogResult = 'Cancel'
        Size         = [System.Drawing.Size]::new(80, 28)
    }
    $cancelButton.Location = [System.Drawing.Point]::new($formWidth - 92, 6)
    $btnPanel.Controls.Add($okButton)
    $btnPanel.Controls.Add($cancelButton)
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton

    $form.Controls.Add($panel)
    $form.Controls.Add($btnPanel)

    if ($form.ShowDialog() -eq 'OK') {
        $selected = $radios | Where-Object { $_.Checked } | Select-Object -First 1
        return [int]$selected.Tag
    }
    return 0  # default / cancel
}

function Invoke-MSBuild {
    param([string]$SolutionFile, [string]$Configuration = 'Release')
    Write-Log "Building: $SolutionFile ($Configuration)"
    $msbuild = 'C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe'
    if (-not (Test-Path $msbuild)) {
        $msbuild = (Get-Command msbuild -ErrorAction SilentlyContinue).Source
    }
    & $msbuild $SolutionFile /p:Configuration=$Configuration /verbosity:minimal
    if ($LASTEXITCODE -ne 0) { throw "MSBuild failed for $SolutionFile" }
}

function Invoke-InstallAware {
    param([string]$ProjectFile)
    Write-Log "Building InstallAware: $ProjectFile"
    & miacmd.exe $ProjectFile
    if ($LASTEXITCODE -ne 0) { throw "InstallAware build failed for $ProjectFile" }
}

function Send-BuildEmail {
    param([string]$To, [string]$From, [string]$Subject, [string]$Body, [string]$SmtpServer, [int]$Port = 25)
    Write-Log "Sending email to $To`: $Subject"
    try {
        $message = [System.Net.Mail.MailMessage]::new($From, $To, $Subject, $Body)
        $smtp = [System.Net.Mail.SmtpClient]::new($SmtpServer, $Port)
        $smtp.Send($message)
        $smtp.Dispose()
        $message.Dispose()
    } catch {
        Write-Log "[WARNING] Failed to send email: $_"
    }
}

$script:BuildLog = [System.Collections.Generic.List[string]]::new()
$script:LogFile = ''
$script:BuildTitle = ''
$script:LAST_STDOUT = ''
$script:LAST_STDERR = ''

function Export-BuildLog {
    param([string]$Path, [string]$Mode = 'Text')
    if ($Path) { $script:BuildLog | Out-File -FilePath $Path -Force }
    return ($script:BuildLog -join "`r`n")
}

#======================================================================
# MAIN SCRIPT
#======================================================================

# TODO: need to cater for missing dcc32 when finding Vault highest version number

# === Submacro: Get files from DevOps ===
function Invoke-Get-files-from-DevOps {
    param(
        $Repository,
        $Output,
        $PASQLFolder,
        $ProjectRoot
    )

    if (Test-Path "$Output") {  # Output
        Remove-ItemSafe -Path "$Output" -Recurse  # Remove Output
    }
    New-Item -ItemType Directory -Path "$Output" -Force | Out-Null  # Create Output
    $PowershellParams = "-OutputFolder '$Output' -Repository '$Repository'"  # PowershellParams
    $CopyScript = "$ProjectRoot\copy_pasql.ps1"  # CopyScript
    # [DISABLED] Log Message
    # [DISABLED] Execute DOS Command (Get DevOps Files) - Get DevOps Files
    Invoke-DosCommand -Command "powershell.exe -Command `"& 'C:\work\BuildStudio\getgithubfiles.ps1' $PowershellParams`""  # Get GitHub Files
    Set-Content -Path "$CopyScript" -Value "$sourceFolder = `"$Output`" $destinationFolder = `"$PASQLFolder`"  Write-Host `"Copying scripts from '$sourceFolder' into '$destinationFolder'`"  # Get the list of files in Folder $destinationFolder $destinationFiles = Get-ChildItem -Path $destinationFolder | ForEach-Object { $_.Name }  # Iterate through each file in Folder $sourceFolder foreach ($file in Get-ChildItem -Path $sourceFolder) {     # Check if the file exists in Folder $destinationFolder     if ($destinationFiles -contains $file.Name) {         # If it exists, copy the file from Folder$sourceFolder to Folder $destinationFolder         Copy-Item -Path $file.FullName -Destination $destinationFolder -Force         Write-Host `"Copied: $($file.Name)`"     } else {         Write-Host `"Skipped: $($file.Name) (Not found in Folder $destinationFolder)`"     } }"  # Generate copy script
    Invoke-DosCommand -Command "powershell.exe -Command `"& '$CopyScript'`""  # Copy files
}


# === Submacro: Update dcc config file ===
function Invoke-Update-dcc-config-file {

    $FILE_COUNT = 0
    foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\Source\dcc*.cfg" -ErrorAction SilentlyContinue)) {  # 
        $CONFIG_FILE = $__file.FullName
        $FILE_COUNT++
        $FILE_NAME_ONLY = Split-Path -Path "$CONFIG_FILE" -Leaf  # Extract file name
        if ("$INI_SECTION" -like "ePay Japanese*") {  # ePay Japanese quirk - has its own dccXX.cfg
            $JAPANESE_FILE_NAME_ONLY = "$FILE_NAME_ONLY"  # Set JAPANESE_FILE_NAME_ONLY
            $JAPANESE_FILE_NAME_ONLY = $JAPANESE_FILE_NAME_ONLY.Replace('.cfg', '-japanese.cfg')  # 
            $JAPANESE_CONFIG_FILE = "$CONFIG_FILE"  # Set JAPANESE_CONFIG_FILE
            $JAPANESE_CONFIG_FILE = $JAPANESE_CONFIG_FILE.Replace('.cfg', '-japanese.cfg')  # 
            Copy-FileEx -Source "$JAPANESE_CONFIG_FILE" -Destination "$CONFIG_FILE" -Force  # Overwrite dccXX.cfg with dccXX-japanese.cfg
            break
        }
        $ORIGINAL_DCC_CFG_FILE_TEXT = Get-Content -Path "$CONFIG_FILE" -Raw  # Read the original dccXXcfg file
        if ("$DELPHI_VERSION" -ceq "6") {  # Delphi 6 TODO: some projects are pegged to a certain version of tools
            Replace-InFile -Path "$CONFIG_FILE" -Find "^\x0D\x0A" -Replace ""  # Update version numbers
        }
        if ("$DELPHI_VERSION" -ceq "2007") {  # Delphi 2007
            Replace-InFile -Path "$CONFIG_FILE" -Find "^\x0D\x0A" -Replace ""  # Update version numbers
        }
        if ("$DELPHI_VERSION" -eq "XE2") {  # Delphi XE2
            Replace-InFile -Path "$CONFIG_FILE" -Find "^\x0D\x0A" -Replace ""  # Update version numbers
        }
        if ("$DELPHI_VERSION" -eq "XE6") {  # Delphi  XE6
            Replace-InFile -Path "$CONFIG_FILE" -Find "^\x0D\x0A" -Replace ""  # Update version numbers
            if ("$CONFIG_FILE" -like "*64.cfg") {  # If using 64bit dcc.cfg
                Replace-InFile -Path "$CONFIG_FILE" -Find "^\x0D\x0A" -Replace ""  # Update Eurekalog 64bit version numbers
            } else {
                Replace-InFile -Path "$CONFIG_FILE" -Find "^\x0D\x0A" -Replace ""  # Update Eurekalog 32bit version numbers
            }
        }
        if ("$DELPHI_VERSION" -eq "10.4") {  # Delphi  10.4
            if ("$CONFIG_FILE" -like "*64.cfg") {  # If using 64bit dcc.cfg
                Replace-InFile -Path "$CONFIG_FILE" -Find "^\x0D\x0A" -Replace ""  # Update 64bit version numbers
            } else {
                Replace-InFile -Path "$CONFIG_FILE" -Find "^\x0D\x0A" -Replace ""  # Update 32bit version numbers
            }
        }
        $UPDATED_DCC_CFG_FILE_TEXT = Get-Content -Path "$CONFIG_FILE" -Raw  # Read the updated dccXXcfg file
        if ("$UPDATED_DCC_CFG_FILE_TEXT" -ne "$ORIGINAL_DCC_CFG_FILE_TEXT") {  # If changes were made to dccXXcfg file
            $VAR_RESULT = Confirm-Action -Message "Local version of the $CONFIG_FILE file was updated. Review the changes below and click `"Yes`" to continue building with the updated file. The updated file will be checked into Vault on successful build.  Otherwise click “No” to roll back the changes and continue the build with the original file.   Updated file:  [$UPDATED_DCC_CFG_FILE_TEXT]  Original file:  [$ORIGINAL_DCC_CFG_FILE_TEXT] " -Default "True"  # 
            if ("$VAR_RESULT" -eq "true") {  # Yes, update
                if ("$DELPHI_VERSION" -ceq "6") {  # Delphi 6
                    $VAR_RESULT = Confirm-Action -Message "This is a Delphi 6 project. Some projects do not compile with the latest available 3rd party packages.  Update $FILE_NAME_ONLY  file for this project? " -Default "False"  # 
                    if ("$VAR_RESULT" -eq "true") {  # Yes, update
                        Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/$FILE_NAME_ONLY" -Host "$VAULT_SERVER_ADDRESS"  # 
                    } else {
                        Invoke-VaultGetLatest -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/$FILE_NAME_ONLY"  # Get Project source (set modification date to Vault)
                    }
                } else {
                    if ("$INI_SECTION" -like "ePay Japanese*") {  # ePay Japanese quirk - has its own dccXX.cfg
                        Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/$JAPANESE_FILE_NAME_ONLY" -Host "$VAULT_SERVER_ADDRESS"  # Check out dccXX-japanese.cfg without overwriting local copy
                        Copy-FileEx -Source "$CONFIG_FILE" -Destination "$JAPANESE_CONFIG_FILE" -Force  # Overwrite local dccXX-japanese.cfg with local dccXX.cfg
                    } else {
                        Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/$FILE_NAME_ONLY" -Host "$VAULT_SERVER_ADDRESS"  # Check out dccXX.cfg without overwriting local copy
                    }
                }
            } else {
                if ("$INI_SECTION" -like "ePay Japanese*") {  # ePay Japanese quirk - has its own dccXX.cfg
                    Copy-FileEx -Source "$JAPANESE_CONFIG_FILE" -Destination "$CONFIG_FILE" -Force  # Overwrite dccXX.cfg with dccXX-japanese.cfg
                } else {
                    Invoke-VaultGetLatest -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/$FILE_NAME_ONLY"  # Overwrite dccXX.cfg with version in VSS
                }
            }
        }
    }
    if ("$FILE_COUNT" -ceq "0") {  # If no files found
        $VAR_RESULT = Confirm-Action -Message "File $CONFIG_FILE does not exist. Continue build?" -Default "True"  # 
        if ("$VAR_RESULT" -ne "true") {
            throw "File $CONFIG_FILE does not exist."
        }
    }
}


# === Submacro: Build project (DPR file) ===
function Invoke-Build-project-DPR-file {
    param(
        $Compiler,
        $ProjectFile
    )

    $ProjectLocation = Split-Path -Path "$ProjectFile" -Parent  # Extract ProjectLocation
    $ProjectName = Split-Path -Path "$ProjectFile" -Leaf  # Extract ProjectName
    $AddEurekaLog = ""  # Clear AddEurekaLog flag
    if ("$ProjectName" -like "*server*") {  # If project name contains server
        $IsServer = "Y"  # IsServer = true
        if (Test-Path "$BUILD_TEMP_PATH\source\EurekaLogOptions?Server.eof") {  # Check for existence of EurekaLogOptions?Server.eof file
            foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\source\EurekaLogOptions?Server.eof" -ErrorAction SilentlyContinue)) {  # 
                $EurekaLogConfigFile = $__file.FullName
                $EurekaLogConfigFile = Split-Path -Path "$EurekaLogConfigFile" -Leaf  # Extract file name
                $AddEurekaLog = "YES"  # Set AddEurekaLog flag
            }
        }
    } else {
        $IsServer = "N"  # IsServer = false
        if (Test-Path "$BUILD_TEMP_PATH\source\EurekaLogOptions?.eof") {  # Check for existence of EurekaLogOptions?.eof file
            foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\source\EurekaLogOptions?.eof" -ErrorAction SilentlyContinue)) {  # 
                $EurekaLogConfigFile = $__file.FullName
                $EurekaLogConfigFile = Split-Path -Path "$EurekaLogConfigFile" -Leaf  # Extract file name
                $AddEurekaLog = "YES"  # Set AddEurekaLog flag
            }
        }
    }
    $ExceptionLogFound = Find-InFile -Path "$ProjectFile" -Find "ExceptionLog"  # Search for "ExceptionLog" in theproject file, result in AddEurekaLog
    if ("$ExceptionLogFound" -ne "") {  # If "ExceptionLog" found
    } else {
        $AddEurekaLog = ""  # Unset AddEurekaLog flag
    }
    if (("$Compiler" -eq "6") -or ("$Compiler" -eq "2007")) {  # If building legacy projects
        # Script block (DelphiScript): Construct project build string into StandardBuild and debug string into DebugBuild
        Write-Log "Constructing compiler command for $Compiler"
    
        # Build the standard (release) command
        $StandardBuild = "dcc32.exe -B -E`"$ProjectLocation..\`" -U`"$DCCUnitSearch`" -D`"$DCCDefines`" `"$ProjectFile`""
    
        # Build the debug command (add -V for debug info)
        $DebugBuild = "dcc32.exe -B -E`"$ProjectLocation..\`" -U`"$DCCUnitSearch`" -D`"$DCCDefines`" -V `"$ProjectFile`""
    
        Write-Log "StandardBuild: $StandardBuild"
        Write-Log "DebugBuild: $DebugBuild"
        
        Invoke-DosCommand -Command "$StandardBuild" -WorkingDirectory "$ProjectLocation"  # Build standard project
        Invoke-DosCommand -Command "$DebugBuild" -WorkingDirectory "$ProjectLocation"  # Build debug project
        Invoke-CheckBuildLogFile   # Check standard build result
        Invoke-CheckBuildLogFile   # Check debug build result
    } else {  # If building XE2 or more recent projects
        $ProjectName = $ProjectName.Replace('.dpr', '.dproj')  # Rename dpr to dproj
        $StandardBuildLog = "${ProjectLocation}Build$ProjectName.log"  # Set StandardBuildLog
        $DebugBuildLog = "${ProjectLocation}Debug-Build$ProjectName.log"  # Set DebugBuildLog
        $Platform = ""  # Set Platform
        if ("$ProjectName" -notlike "*X.dproj") {  # If not building activeX
            # If XML Node/Attribute Exists: //PropertyGroup[1]/Platform[1]/text()[1]
            $PlatformFound = 'False'
            try {
                [xml]$_xml = Get-Content -Path "$ProjectLocation$ProjectName" -Raw
                $_nsMgr = [System.Xml.XmlNamespaceManager]::new($_xml.NameTable)
                $_node = $_xml.SelectSingleNode("//*[local-name()='Platform'][@*[local-name()='Condition']] ", $_nsMgr)
                if ($_node) { $PlatformFound = 'True' }
            } catch { Write-Log "XML query failed: $_" }
            if ($PlatformFound -ceq 'True') {
                $None = Get-XmlValue -Path "" -XPath "//*[local-name()='Platform'][@*[local-name()='Condition']] "  # Platform
            } else {
            }
        } else {
            $Platform = "Win32"  # Platform = Win32
        }

        #region Generate build scripts
        Write-Log "--- Generate build scripts ---"
        $DCCDefines = ""  # Clear DCCDefines
        $DCCUnitSearch = ""  # Clear DCCUnitSearch
        $DCCConsoleTarget = ""  # Clear DCCConsoleTarget
        if ("$Platform" -eq "Win64") {  # if Platform == win64
            if (Test-Path "$ProjectLocation\dcc64.cfg") {  # If dcc64.cfg file exists
                foreach ($CurrentLine in (Get-Content -Path "$ProjectLocation\dcc64.cfg")) {  # Loop through dcc64.cfg
                    $CurrentLine = ("$CurrentLine").Trim()  # Trim CurrentLine
                    if (("$CurrentLine").StartsWith("-D")) {  # If CurrentLine starts with -D
                        # [DISABLED] String Substring (Extract directives) - Extract directives
                        $CurrentLine = $CurrentLine.Replace('-D', ' ')  # Extract directives
                        $CurrentLine = ("$CurrentLine").Trim()  # Trim CurrentLine
                        $DCCDefines = "$DCCDefines$CurrentLine;"  # Set DCCDefines
                    }
                    if (("$CurrentLine").StartsWith("-u")) {  # If CurrentLine starts with -u
                        # [DISABLED] String Substring (Extract path) - Extract path
                        $CurrentLine = $CurrentLine.Replace('-u', ' ')  # Extract path
                        $CurrentLine = $CurrentLine.Replace('"', ' ')  # Extract path
                        $CurrentLine = ("$CurrentLine").Trim()  # Trim CurrentLine
                        $DCCUnitSearch = "$DCCUnitSearch$CurrentLine;"  # Set DCCUnitSearch
                    }
                }
            } else {
                throw "Unable to locate configuration file `"$ProjectLocation\dcc64.cfg`". Please make sure this file exists in the project source folder."
            }
        } else {
            if ("$ProjectName" -notlike "*X.dproj") {  # If not building activeX
                if (Test-Path "${ProjectLocation}dcc32.cfg") {  # If dcc32.cfg file exists
                    foreach ($CurrentLine in (Get-Content -Path "$ProjectLocation\dcc32.cfg")) {  # Loop through dcc32.cfg
                        $CurrentLine = ("$CurrentLine").Trim()  # Trim CurrentLine
                        if (("$CurrentLine").StartsWith("-D")) {  # If CurrentLine starts with -D
                            # [DISABLED] String Substring (Extract directives) - Extract directives
                            $CurrentLine = $CurrentLine.Replace('-D', ' ')  # Extract directives
                            $CurrentLine = ("$CurrentLine").Trim()  # Trim CurrentLine
                            $DCCDefines = "$DCCDefines$CurrentLine;"  # Set DCCDefines
                        }
                        if (("$CurrentLine").StartsWith("-u")) {  # If CurrentLine starts with -u
                            # [DISABLED] String Substring (Extract path) - Extract path
                            $CurrentLine = $CurrentLine.Replace('-u', ' ')  # Extract path
                            $CurrentLine = $CurrentLine.Replace('"', ' ')  # Extract path
                            $CurrentLine = ("$CurrentLine").Trim()  # Trim CurrentLine
                            $DCCUnitSearch = "$DCCUnitSearch$CurrentLine;"  # Set DCCUnitSearch
                        }
                    }
                } else {
                    throw "Unable to locate configuration file `"$ProjectLocation\dcc32.cfg`". Please make sure this file exists in the project source folder."
                }
            }
        }
        $DCCUnitSearch = "$DCCUnitSearch$COMPILER_EUREKALOG_COMMON_PATH;"  # Set DCCUnitSearch for Erekalog common files
        $DCCDefines = ("$DCCDefines").Trim()  # Trim DCCDefines
        $DCCUnitSearch = ("$DCCUnitSearch").Trim()  # Trim DCCUnitSearch
        if ("$ProjectName" -like "*tests.dproj") {  # If project ends with TESTS.dproj
            $DCCConsoleTarget = ";DCC_ConsoleTarget=true"  # Set DCCConsoleTarget compiler directive
            # [DISABLED] Set/Reset Variable Value (Set DCCConsoleTarget compiler directive) - Set DCCConsoleTarget compiler directive
        }
        if ("$ProjectName" -like "*console.dproj") {  # If project ends with CONSOLE.dproj
            $DCCConsoleTarget = ";DCC_ConsoleTarget=true"  # Set DCCConsoleTarget compiler directive
        }
        if (("$ProjectName" -like "*PAUnitCMD.dpr") -and ("$ProjectName" -like "*PAUnitCMD.dproj")) {  # PAUnit quirk - If project ends with PAUnitCMD.dproj
            $DCCConsoleTarget = ";DCC_ConsoleTarget=true"  # Set DCCConsoleTarget compiler directive
        }
        if ("$DCCDefines" -ne "") {  # If DCCDefines not blank
            $DCCDefines = "@SET DCC_Define=$DCCDefines"  # Prepend DCCDefines
        }
        if ("$DCCUnitSearch" -ne "") {  # If DCCUnitSearch not blank
            $DCCUnitSearch = "@SET DCC_UnitSearchPath=$DCCUnitSearch"  # Prepend DCCUnitSearch
        }
        if ("$Compiler" -like "*10.4*") {  # if 10.4
            if ("$Platform" -eq "Win64") {  # if Platform == win64
                if (Test-Path "$ProjectLocation\dcc64.cfg") {  # If dcc64.cfg file exists
                    Set-Content -Path "${ProjectLocation}rsvars.bat" -Value "@SET BDS=C:\Compilers\Delphi 10.4\Embarcadero @SET BDSINCLUDE=C:\Compilers\Delphi 10.4\Embarcadero\include @SET BDSCOMMONDIR=C:\Users\Public\Documents\Embarcadero @SET FrameworkDir=C:\Windows\Microsoft.NET\Framework\v4.0.30319 @SET FrameworkVersion=v4.0 @SET FrameworkSDKDir= @SET PATH=$FrameworkDir;$FrameworkSDKDir;C:\Compilers\Delphi 10.4\Embarcadero\bin;C:\Compilers\Delphi 10.4\Embarcadero\bin64;$PATH @SET LANGDIR=EN @SET PLATFORM= @SET PlatformSDK= $DCCDefines $DCCUnitSearch;C:\Compilers\Delphi 10.4\Embarcadero\lib\Win64\debug;C:\Compilers\Delphi 10.4\Embarcadero\Source\DUnit\src;C:\Compilers\Delphi 10.4\Embarcadero\lib\Win64\release;C:\Compilers\Delphi 10.4\Embarcadero\Imports;C:\Compilers\Delphi 10.4\Embarcadero\include @SET DCC_UsePackage=dxTileControlRS16;dxdborRS16;dxPScxVGridLnkRS16;cxLibraryRS16;dxLayoutControlRS16;dxPScxPivotGridLnkRS16;dxCoreRS16;cxExportRS16;dxBarRS16;cxSpreadSheetRS16;cxTreeListdxBarPopupMenuRS16;TeeDB;dxDBXServerModeRS16;dxPsPrVwAdvRS16;dxPSCoreRS16;dxPScxTLLnkRS16;dxPScxGridLnkRS16;cxPageControlRS16;dxRibbonRS16;DBXSybaseASEDriver;vclimg;cxTreeListRS16;dxComnRS16;vcldb;vcldsnap;dxBarExtDBItemsRS16;DBXDb2Driver;vcl;DBXMSSQLDriver;cxDataRS16;cxBarEditItemRS16;dxDockingRS16;dxPSDBTeeChartRS16;cxPageControldxBarPopupMenuRS16;cxSchedulerGridRS16;dxBarExtItemsRS16;dxPSLnksRS16;dxtrmdRS16;adortl;dxPSTeeChartRS16;cxVerticalGridRS16;dxPSdxLCLnkRS16;dxorgcRS16;dxWizardControlRS16;dxPScxExtCommonRS16;dxNavBarRS16;dxPSdxDBOCLnkRS16;cxSchedulerTreeBrowserRS16;Tee;DBXOdbcDriver;dxdbtrRS16;dxPScxCommonRS16;dxmdsRS16;dxSpellCheckerRS16;dxPScxSSLnkRS16;cxGridRS16;dxPSPrVwRibbonRS16;cxEditorsRS16;vclactnband;TeeUI;bindcompvcl;dxServerModeRS16;cxPivotGridRS16;dxPScxSchedulerLnkRS16;vclie;cxSchedulerRS16;vcltouch;cxSchedulerRibbonStyleEventEditorRS16;dxPSdxDBTVLnkRS16;VclSmp;dxTabbedMDIRS16;dxPSdxOCLnkRS16;dsnapcon;dxPSdxFCLnkRS16;dxThemeRS16;dxPScxPCProdRS16;vclx;dxFlowChartRS16;dxGDIPlusRS16;dxBarDBNavRS16;PGPTLSBBoxD16;fmx;IndySystem;DCBBoxD16;CloudBBoxD16;DBXInterBaseDriver;OfficeBBoxD16;DataSnapCommon;DbxCommonDriver;EDIBBoxD16;dbxcds;FTPSBBoxCliD16;ZIPBBoxD16;DBXOracleDriver;CustomIPTransport;HTTPBBoxCliD16;dsnap;IndyCore;HTTPBBoxSrvD16;FmxTeeUI;DAVBBoxCliD16;inetdbxpress;SSLBBoxCliD16;PGPMIMEBBoxD16;XMLBBoxD16;IPIndyImpl;SSHBBoxCliD16;LDAPBBoxD16;bindcompfmx;XMLBBoxSecD16;rtl;dbrtl;DbxClientDriver;bindcomp;DsgnBBoxD16;xmlrtl;PDFBBoxD16;IndyProtocols;FTPSBBoxSrvD16;FMXTee;bindengine;DAVBBoxSrvD16;PGPBBoxD16;SSLBBoxSrvD16;MailBBoxD16;MIMEBBoxD16;SMIMEBBoxD16;DBXInformixDriver;PGPSSHBBoxD16;PGPLDAPBBoxD16;BaseBBoxD16;DBXFirebirdDriver;inet;DBXSybaseASADriver;dbexpress; @SET DCC_MapFile=3 "  # Create rsvars.bat file
                    Set-Content -Path "${ProjectLocation}build.bat" -Value "call `"${ProjectLocation}rsvars.bat`" rem build release msbuild `"$ProjectName`" /verbosity:diag /clp:ShowCommandLine /t:Rebuild /p:Config=Release;platform=Win64;EnvOptionsWarn=false$DCCConsoleTarget;DCC_ExeOutput=..\;DCC_LocalDebugSymbols=false;DCC_SymbolReferenceInfo=0;DCC_DebugInformation=2 > `"$StandardBuildLog`" "  # Create build.bat file
                    Set-Content -Path "${ProjectLocation}debug-build.bat" -Value "call `"${ProjectLocation}rsvars.bat`" rem build debug msbuild `"$ProjectName`" /verbosity:diag /clp:ShowCommandLine /t:Rebuild /p:Config=Debug;platform=Win64;EnvOptionsWarn=false$DCCConsoleTarget;DCC_PlatformTarget=Win64;DCC_RemoteDebug=true;DCC_ExeOutput=..\debug;DCC_MapFile=3;DCC_DebugInfoInExe=true;DCC_Optimize=false;DCC_DebugDCUs=true;DCC_GenerateStackFrames=true > `"$DebugBuildLog`""  # Create debug-build.bat file
                } else {
                    throw "Unable to locate configuration file `"$ProjectLocation\dcc64.cfg`". Please make sure this file exists in the project source folder."
                }
            } else {
                Set-Content -Path "${ProjectLocation}rsvars.bat" -Value "@SET BDS=C:\Compilers\Delphi 10.4\Embarcadero @SET BDSINCLUDE=C:\Compilers\Delphi 10.4\Embarcadero\include @SET BDSCOMMONDIR=C:\Users\Public\Documents\Embarcadero @SET FrameworkDir=C:\Windows\Microsoft.NET\Framework\v4.0.30319 @SET FrameworkVersion=v4.0 @SET FrameworkSDKDir= @SET PATH=$FrameworkDir;$FrameworkSDKDir;C:\Compilers\Delphi 10.4\Embarcadero\bin;C:\Compilers\Delphi 10.4\Embarcadero\bin64;$PATH @SET LANGDIR=EN @SET PLATFORM= @SET PlatformSDK= $DCCDefines $DCCUnitSearch;C:\Compilers\Delphi 10.4\Embarcadero\lib\Win32\debug;C:\Compilers\Delphi 10.4\Embarcadero\Source\DUnit\src;C:\Compilers\Delphi 10.4\Embarcadero\lib\Win32\release;C:\Compilers\Delphi 10.4\Embarcadero\Imports;C:\Compilers\Delphi 10.4\Embarcadero\include @SET DCC_UsePackage=dxTileControlRS16;dxdborRS16;dxPScxVGridLnkRS16;cxLibraryRS16;dxLayoutControlRS16;dxPScxPivotGridLnkRS16;dxCoreRS16;cxExportRS16;dxBarRS16;cxSpreadSheetRS16;cxTreeListdxBarPopupMenuRS16;TeeDB;dxDBXServerModeRS16;dxPsPrVwAdvRS16;dxPSCoreRS16;dxPScxTLLnkRS16;dxPScxGridLnkRS16;cxPageControlRS16;dxRibbonRS16;DBXSybaseASEDriver;vclimg;cxTreeListRS16;dxComnRS16;vcldb;vcldsnap;dxBarExtDBItemsRS16;DBXDb2Driver;vcl;DBXMSSQLDriver;cxDataRS16;cxBarEditItemRS16;dxDockingRS16;dxPSDBTeeChartRS16;cxPageControldxBarPopupMenuRS16;cxSchedulerGridRS16;dxBarExtItemsRS16;dxPSLnksRS16;dxtrmdRS16;adortl;dxPSTeeChartRS16;cxVerticalGridRS16;dxPSdxLCLnkRS16;dxorgcRS16;dxWizardControlRS16;dxPScxExtCommonRS16;dxNavBarRS16;dxPSdxDBOCLnkRS16;cxSchedulerTreeBrowserRS16;Tee;DBXOdbcDriver;dxdbtrRS16;dxPScxCommonRS16;dxmdsRS16;dxSpellCheckerRS16;dxPScxSSLnkRS16;cxGridRS16;dxPSPrVwRibbonRS16;cxEditorsRS16;vclactnband;TeeUI;bindcompvcl;dxServerModeRS16;cxPivotGridRS16;dxPScxSchedulerLnkRS16;vclie;cxSchedulerRS16;vcltouch;cxSchedulerRibbonStyleEventEditorRS16;dxPSdxDBTVLnkRS16;VclSmp;dxTabbedMDIRS16;dxPSdxOCLnkRS16;dsnapcon;dxPSdxFCLnkRS16;dxThemeRS16;dxPScxPCProdRS16;vclx;dxFlowChartRS16;dxGDIPlusRS16;dxBarDBNavRS16;PGPTLSBBoxD16;fmx;IndySystem;DCBBoxD16;CloudBBoxD16;DBXInterBaseDriver;OfficeBBoxD16;DataSnapCommon;DbxCommonDriver;EDIBBoxD16;dbxcds;FTPSBBoxCliD16;ZIPBBoxD16;DBXOracleDriver;CustomIPTransport;HTTPBBoxCliD16;dsnap;IndyCore;HTTPBBoxSrvD16;FmxTeeUI;DAVBBoxCliD16;inetdbxpress;SSLBBoxCliD16;PGPMIMEBBoxD16;XMLBBoxD16;IPIndyImpl;SSHBBoxCliD16;LDAPBBoxD16;bindcompfmx;XMLBBoxSecD16;rtl;dbrtl;DbxClientDriver;bindcomp;DsgnBBoxD16;xmlrtl;PDFBBoxD16;IndyProtocols;FTPSBBoxSrvD16;FMXTee;bindengine;DAVBBoxSrvD16;PGPBBoxD16;SSLBBoxSrvD16;MailBBoxD16;MIMEBBoxD16;SMIMEBBoxD16;DBXInformixDriver;PGPSSHBBoxD16;PGPLDAPBBoxD16;BaseBBoxD16;DBXFirebirdDriver;inet;DBXSybaseASADriver;dbexpress; @SET DCC_MapFile=3"  # Create rsvars.bat file
                Set-Content -Path "${ProjectLocation}build.bat" -Value "call `"${ProjectLocation}rsvars.bat`" rem build release msbuild `"$ProjectName`" /verbosity:diag /clp:ShowCommandLine /t:Rebuild /p:Config=Release;platform=Win32;EnvOptionsWarn=false$DCCConsoleTarget;DCC_ExeOutput=..\;DCC_LocalDebugSymbols=false;DCC_SymbolReferenceInfo=0;DCC_DebugInformation=2 > `"$StandardBuildLog`" "  # Create build.bat file
                Set-Content -Path "${ProjectLocation}debug-build.bat" -Value "call `"${ProjectLocation}rsvars.bat`" rem build debug msbuild `"$ProjectName`" /verbosity:diag /clp:ShowCommandLine /t:Rebuild /p:Config=Debug;platform=Win32;EnvOptionsWarn=false$DCCConsoleTarget;DCC_PlatformTarget=Win32;DCC_RemoteDebug=true;DCC_ExeOutput=..\debug;DCC_MapFile=3;DCC_DebugInfoInExe=true;DCC_Optimize=false;DCC_DebugDCUs=true;DCC_GenerateStackFrames=true > `"$DebugBuildLog`""  # Create debug-build.bat file
            }
        } else {
            if ("$Compiler" -like "*XE6*") {  # if XE6
                if (Test-Path "$ProjectLocation\dcc64.cfg") {  # If dcc64.cfg file exists
                    Set-Content -Path "${ProjectLocation}rsvars.bat" -Value "@SET BDS=C:\Compilers\Delphi XE6\Embarcadero @SET BDSINCLUDE=C:\Compilers\Delphi XE6\Embarcadero\include @SET BDSCOMMONDIR=C:\Users\Public\Documents\Embarcadero @SET FrameworkDir=C:\Windows\Microsoft.NET\Framework\v3.5 @SET FrameworkVersion=v3.5 @SET FrameworkSDKDir= @SET PATH=$FrameworkDir;$FrameworkSDKDir;C:\Compilers\Delphi XE6\Embarcadero\bin;C:\Compilers\Delphi XE6\Embarcadero\bin64;$PATH @SET LANGDIR=EN @SET PLATFORM= @SET PlatformSDK= $DCCDefines $DCCUnitSearch;C:\Compilers\Delphi XE6\Embarcadero\lib\Win64\debug;C:\Compilers\Delphi XE6\Embarcadero\Source\DUnit\src;C:\Compilers\Delphi XE6\Embarcadero\lib\Win64\release;C:\Compilers\Delphi XE6\Embarcadero\Imports;C:\Compilers\Delphi XE6\Embarcadero\include @SET DCC_UsePackage=dxTileControlRS16;dxdborRS16;dxPScxVGridLnkRS16;cxLibraryRS16;dxLayoutControlRS16;dxPScxPivotGridLnkRS16;dxCoreRS16;cxExportRS16;dxBarRS16;cxSpreadSheetRS16;cxTreeListdxBarPopupMenuRS16;TeeDB;dxDBXServerModeRS16;dxPsPrVwAdvRS16;dxPSCoreRS16;dxPScxTLLnkRS16;dxPScxGridLnkRS16;cxPageControlRS16;dxRibbonRS16;DBXSybaseASEDriver;vclimg;cxTreeListRS16;dxComnRS16;vcldb;vcldsnap;dxBarExtDBItemsRS16;DBXDb2Driver;vcl;DBXMSSQLDriver;cxDataRS16;cxBarEditItemRS16;dxDockingRS16;dxPSDBTeeChartRS16;cxPageControldxBarPopupMenuRS16;cxSchedulerGridRS16;dxBarExtItemsRS16;dxPSLnksRS16;dxtrmdRS16;adortl;dxPSTeeChartRS16;cxVerticalGridRS16;dxPSdxLCLnkRS16;dxorgcRS16;dxWizardControlRS16;dxPScxExtCommonRS16;dxNavBarRS16;dxPSdxDBOCLnkRS16;cxSchedulerTreeBrowserRS16;Tee;DBXOdbcDriver;dxdbtrRS16;dxPScxCommonRS16;dxmdsRS16;dxSpellCheckerRS16;dxPScxSSLnkRS16;cxGridRS16;dxPSPrVwRibbonRS16;cxEditorsRS16;vclactnband;TeeUI;bindcompvcl;dxServerModeRS16;cxPivotGridRS16;dxPScxSchedulerLnkRS16;vclie;cxSchedulerRS16;vcltouch;cxSchedulerRibbonStyleEventEditorRS16;dxPSdxDBTVLnkRS16;VclSmp;dxTabbedMDIRS16;dxPSdxOCLnkRS16;dsnapcon;dxPSdxFCLnkRS16;dxThemeRS16;dxPScxPCProdRS16;vclx;dxFlowChartRS16;dxGDIPlusRS16;dxBarDBNavRS16;PGPTLSBBoxD16;fmx;IndySystem;DCBBoxD16;CloudBBoxD16;DBXInterBaseDriver;OfficeBBoxD16;DataSnapCommon;DbxCommonDriver;EDIBBoxD16;dbxcds;FTPSBBoxCliD16;ZIPBBoxD16;DBXOracleDriver;CustomIPTransport;HTTPBBoxCliD16;dsnap;IndyCore;HTTPBBoxSrvD16;FmxTeeUI;DAVBBoxCliD16;inetdbxpress;SSLBBoxCliD16;PGPMIMEBBoxD16;XMLBBoxD16;IPIndyImpl;SSHBBoxCliD16;LDAPBBoxD16;bindcompfmx;XMLBBoxSecD16;rtl;dbrtl;DbxClientDriver;bindcomp;DsgnBBoxD16;xmlrtl;PDFBBoxD16;IndyProtocols;FTPSBBoxSrvD16;FMXTee;bindengine;DAVBBoxSrvD16;PGPBBoxD16;SSLBBoxSrvD16;MailBBoxD16;MIMEBBoxD16;SMIMEBBoxD16;DBXInformixDriver;PGPSSHBBoxD16;PGPLDAPBBoxD16;BaseBBoxD16;DBXFirebirdDriver;inet;DBXSybaseASADriver;dbexpress; @SET DCC_MapFile=3"  # Create rsvars.bat file
                    Set-Content -Path "${ProjectLocation}build.bat" -Value "call `"${ProjectLocation}rsvars.bat`" rem build release msbuild `"$ProjectName`" /verbosity:diag /clp:ShowCommandLine /t:Rebuild /p:Config=Release;platform=Win64;EnvOptionsWarn=false$DCCConsoleTarget;DCC_ExeOutput=..\;DCC_LocalDebugSymbols=false;DCC_SymbolReferenceInfo=0;DCC_DebugInformation=2 > `"$StandardBuildLog`" "  # Create build.bat file
                    Set-Content -Path "${ProjectLocation}debug-build.bat" -Value "call `"${ProjectLocation}rsvars.bat`" rem build debug msbuild `"$ProjectName`" /verbosity:diag /clp:ShowCommandLine /t:Rebuild /p:Config=Debug;platform=Win64;EnvOptionsWarn=false$DCCConsoleTarget;DCC_PlatformTarget=Win64;DCC_RemoteDebug=true;DCC_ExeOutput=..\debug;DCC_MapFile=3;DCC_DebugInfoInExe=true;DCC_Optimize=false;DCC_DebugDCUs=true;DCC_GenerateStackFrames=true > `"$DebugBuildLog`""  # Create debug-build.bat file
                } else {
                    Set-Content -Path "${ProjectLocation}rsvars.bat" -Value "@SET BDS=C:\Compilers\Delphi $Compiler\Embarcadero @SET BDSCOMMONDIR=C:\Users\Public\Documents\RAD Studio\9.0 @SET FrameworkDir=C:\Windows\Microsoft.NET\Framework\v3.5 @SET FrameworkVersion=v3.5 @SET FrameworkSDKDir= @SET PATH=$FrameworkDir;$FrameworkSDKDir;$BDS\bin;$PATH @SET LANGDIR=EN @SET DCC_MapFile=3 $DCCDefines $DCCUnitSearch"  # Create rsvars.bat file
                    Set-Content -Path "${ProjectLocation}build.bat" -Value "call `"${ProjectLocation}rsvars.bat`" rem build release msbuild `"$ProjectName`" /verbosity:diag /clp:ShowCommandLine /t:Rebuild /p:Config=Release;platform=Win32;EnvOptionsWarn=false$DCCConsoleTarget;DCC_ExeOutput=..\;DCC_LocalDebugSymbols=false;DCC_SymbolReferenceInfo=0;DCC_DebugInformation=2 > `"$StandardBuildLog`" "  # Create build.bat file
                    Set-Content -Path "${ProjectLocation}debug-build.bat" -Value "call `"${ProjectLocation}rsvars.bat`" rem build debug msbuild `"$ProjectName`" /verbosity:diag /clp:ShowCommandLine /t:Rebuild /p:Config=Release;platform=Win32;EnvOptionsWarn=false$DCCConsoleTarget;DCC_ExeOutput=..\debug;DCC_DebugInfoInExe=true;DCC_Optimize=false;DCC_GenerateStackFrames=true > `"$DebugBuildLog`" "  # Create debug-build.bat file
                }
            } else {
                Set-Content -Path "${ProjectLocation}rsvars.bat" -Value "@SET BDS=C:\Compilers\Delphi $Compiler\Embarcadero @SET BDSCOMMONDIR=C:\Users\Public\Documents\RAD Studio\9.0 @SET FrameworkDir=C:\Windows\Microsoft.NET\Framework\v3.5 @SET FrameworkVersion=v3.5 @SET FrameworkSDKDir= @SET PATH=$FrameworkDir;$FrameworkSDKDir;$BDS\bin;$PATH @SET LANGDIR=EN @SET DCC_MapFile=3 $DCCDefines $DCCUnitSearch"  # Create rsvars.bat file
                Set-Content -Path "${ProjectLocation}build.bat" -Value "call `"${ProjectLocation}rsvars.bat`" rem build release msbuild `"$ProjectName`" /verbosity:diag /clp:ShowCommandLine /t:Rebuild /p:Config=Release;platform=Win32;EnvOptionsWarn=false$DCCConsoleTarget;DCC_ExeOutput=..\;DCC_LocalDebugSymbols=false;DCC_SymbolReferenceInfo=0;DCC_DebugInformation=true > `"$StandardBuildLog`" "  # Create build.bat file
                Set-Content -Path "${ProjectLocation}debug-build.bat" -Value "call `"${ProjectLocation}rsvars.bat`" rem build debug msbuild `"$ProjectName`" /verbosity:diag /clp:ShowCommandLine /t:Rebuild /p:Config=Release;platform=Win32;EnvOptionsWarn=false$DCCConsoleTarget;DCC_ExeOutput=..\debug;DCC_DebugInfoInExe=true;DCC_Optimize=false;DCC_GenerateStackFrames=true > `"$DebugBuildLog`" "  # Create debug-build.bat file
            }
        }
        #endregion Generate build scripts

        try {  # 
            Invoke-DosCommand -Command "`"${ProjectLocation}build.bat`"" -WorkingDirectory "$ProjectLocation"  # Build standard projects
        } finally {  # 
            Invoke-CheckBuildLogFile   # Check standard build result
        }
        try {  # 
            Invoke-DosCommand -Command "`"${ProjectLocation}debug-build.bat`"" -WorkingDirectory "$ProjectLocation"  # Build debug projects
        } finally {  # 
            Invoke-CheckBuildLogFile   # Check debug build result
        }
        $ExeName = "$ProjectName"  # Set ExeName
        $ExeName = $ExeName.Replace('.dproj', '.exe')  # Replace .dproj with .exe in ExeName
        if ("$AddEurekaLog" -ne "") {  # If Eureka Log is to be added
            $EurekaLogIDEVersion = ""  # Initialise EurekaLogIDEVersion to blank
            # Switch: 
            switch ($VAR_RESULT) {
                "XE2" {  # XE2
                    $EurekaLogIDEVersion = "16"  # 16
                }
                "XE6" {  # XE6
                    $EurekaLogIDEVersion = "20"  # 20
                }
                "10.4" {  # 10.4
                    $EurekaLogIDEVersion = "27"  # 27
                }
            }
            if ("$EurekaLogIDEVersion" -ceq "") {  # if EurekaLogIDEVersion is blank
                throw "Unknown IDE version for compiler `"$Compiler`""
            }
            Invoke-DosCommand -Command "$EUREKALOG_PATH --el_alter_exe`"$ProjectLocation$ProjectName;$ProjectLocation..\$ExeName`" --el_verbose --el_config`"$ProjectLocation$EurekaLogConfigFile`" --el_mode=Delphi --el_profile=Release --el_UnicodeOutput --el_ide=$EurekaLogIDEVersion > `"${ProjectLocation}Eureka-$ExeName.log`"" -WorkingDirectory "$ProjectLocation"  # Inject EurekaLog to the standard project
            # Check result in log file
        }
        if (Test-Path "$ProjectLocation..\$ExeName") {  # If exe exists
            Invoke-SignFile   # Sign exe
        }
        # [DISABLED] Stop Macro Execution
    }
}


# === Submacro: Check build log file ===
function Invoke-Check-build-log-file {
    param(
        $LogFile,
        $Compiler
    )

    if ($Compiler -eq "XE2" -or $Compiler -eq "XE6" -or $Compiler -eq "10.4") {  # If XE2, XE6, 10.4
        $LogText = Find-InFile -Path "$LogFile" -Find "Build succeeded."  # Search log file for "Build succeeded." string
        if ("$LogText" -ne "1") {  # If not found
            $LogText = Find-InFile -Path "$LogFile" -Find "Build FAILED."  # Search log file for "Build FAILED." string
            if ("$LogText" -ne "1") {  # If not found
                throw "Searching log file `"$LogFile`" for either `"Build succeeded.`" or `"Build FAILED.`" produced no result. "
            } else {
                $ResultText = ""  # Clear ResultText
                $LogText = Get-Content -Path "$LogFile" -Raw  # Read LogFile into LogText
                $ResultText = Get-SubstringBetween -Input "$LogText" -Start "Task `"DCC`"" -End "Done executing task `"DCC`""  # Extract error message
                if ("$ResultText" -ne "") {  # If ResultText not empty
                    throw "Failed to build project:  $ResultText"
                }
            }
        }
    } else {
        $LogText = Get-Content -Path "$LogFile" -Raw  # Load log file text
        # Script block (DelphiScript): Replace CR with CRLF in text
        $LogText = $LogText -replace "`r(?!`n)", "`r`n"
        Set-Content -Path "$LogFile" -Value "$LogText"  # Write text back to log file
        $LogText = Invoke-DosCommand -Command "`"C:\Compilers\RAD Studio 2007\CodeGear\Bin\grep`" -i fatal: `"$LogFile`""  # Check for errors
        $LogText = $LogText.Replace('"C:\Compilers\RAD Studio 2007\CodeGear\Bin\grep" -i fatal: "$LogFile"', '.')  # Remove command line from string
        $LogText = $LogText.Replace('cmd /C "."', ' ')  # Remove command line from string
        if (("$LogText" -like "*error:*") -and ("$LogText" -like "*fatal:*")) {  # Check for "error:" "fatal:"
            throw "Failed to build project:  $LogText"
        }
    }
}


# === Submacro: Sign file ===
function Invoke-Sign-file {
    param(
        $FileName,
        $Description
    )

    # [DISABLED] Execute DOS Command (Sign final file) - Sign final file
    $CommandResultText = Invoke-DosCommand -Command "powershell.exe -Command `"& 'C:\work\BuildStudio\signfile.ps1' -FilePath '$FileName' -Description '$Description'`""  # Sign file
    if ("$CommandResult" -ne "0") {  # Check for errors
        throw "Failed to sign setup EXE  $CommandResultText"
    }
    $CommandResultText = Invoke-DosCommand -Command "C:\Compilers\SignTool\signtool.exe verify /pa `"$FileName`""  # Verify file signature
    if ("$CommandResult" -ne "0") {  # Check for errors
        throw "Failed to verify EXE file:  $CommandResultText"
    }
}

try {  # 

    #region Build process
    Write-Log "--- Build process ---"
    # LABEL:   (GoTo not supported in PS - restructure logic)
    if (("$NIGHTLY_BUILD" -eq "TRUE") -and ("$INI_SECTION" -eq "")) {  # If nightly build check if project is set
        throw "No project selected for nightly build"
    }

    #region Read settings
    Write-Log "--- Read settings ---"
    # Script block (DelphiScript): Get current year (for copyright etc.)
    $CURRENT_YEAR = (Get-Date).Year
    $BUILD_YEAR = $CURRENT_YEAR  # BUILD_YEAR is used for major version bumps
    Write-Log "Current year: $CURRENT_YEAR"
    Write-Log "[DEBUG] Checking INI file: ${ABSOPENEDPROJECTDIR}PAApplications.ini exists = $(Test-Path "${ABSOPENEDPROJECTDIR}PAApplications.ini")"
    if (Test-Path "${ABSOPENEDPROJECTDIR}PAApplications.ini") {  # Check for PAApplications.ini
        Write-Log "[DEBUG] NIGHTLY_BUILD='$NIGHTLY_BUILD' NEXT_PROJECT_TO_BUILD='$NEXT_PROJECT_TO_BUILD'"
        if ("$NIGHTLY_BUILD" -ne "TRUE") {  # If not nightly build
            Write-Log "[DEBUG] Not nightly build, checking NEXT_PROJECT_TO_BUILD"
            if ("$NEXT_PROJECT_TO_BUILD" -ne "0") {  # If building another project in a loop
                Write-Log "[DEBUG] Using NEXT_PROJECT_TO_BUILD: $NEXT_PROJECT_TO_BUILD"
                $VAR_RESULT = "$NEXT_PROJECT_TO_BUILD"  # Set VAR_RESULT to NEXT_PROJECT_TO_BUILD
            } else {
                Write-Log "[DEBUG] About to read INI sections and show radio dialog..."
                # Radio Group: Select project to build (dynamic from INI file)
                $iniSections = [System.Collections.Generic.List[string]]::new()
                foreach ($line in [System.IO.File]::ReadLines("${ABSOPENEDPROJECTDIR}PAApplications.ini")) {
                    if ($line -match '^\[(.+)\]$') { $iniSections.Add($Matches[1]) }
                }
                Write-Log "[DEBUG] Found $($iniSections.Count) INI sections"
                $projectOptions = @("None - cancel build") + $iniSections.ToArray()
                Write-Log "[DEBUG] Calling Show-RadioMenu now..."
                $VAR_RESULT = Show-RadioMenu -Title "Select project to build" -Options $projectOptions
                Write-Log "[DEBUG] Show-RadioMenu returned: $VAR_RESULT"
                if ($VAR_RESULT -eq 0) {  # Cancel build
                    return
                }
                $INI_SECTION = $projectOptions[$VAR_RESULT]
            }
            if ("$VAR_RESULT" -ceq "0") {  # Cancel build
                return  # Stop Macro Execution - 
            }
            # Switch:  (skipped — project selection handled by INI-based radio group above)
        }
        Write-Log "[DEBUG] INI_SECTION='$INI_SECTION', reading INI values..."
        $PROJECT_TITLE = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "PROJECT_TITLE"  # PROJECT_TITLE
        Write-Log "[DEBUG] PROJECT_TITLE='$PROJECT_TITLE'"
        $SOURCE_CONTROL_ROOT_PATH = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "SOURCE_CONTROL_ROOT_PATH"  # SOURCE_CONTROL_ROOT_PATH
        $SOURCE_CONTROL_SOURCE_PATH = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "SOURCE_CONTROL_SOURCE_PATH"  # SOURCE_CONTROL_SOURCE_PATH
        $PROJECT_LABEL_PATH = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "SOURCE_CONTROL_SOURCE_PATH"  # PROJECT_LABEL_PATH
        $SOURCE_CONTROL_SOURCE_PATH = ("$SOURCE_CONTROL_SOURCE_PATH").Trim()  # 
        $HELP_FILE_LOCATION = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "HELP_FILE_LOCATION"  # HELP_FILE_LOCATION
        Write-Log "[DEBUG] SOURCE_CONTROL_SOURCE_PATH='$SOURCE_CONTROL_SOURCE_PATH'"
        $DELPHI_VERSION = ""  # Initialise Delphi compiler version DELPHI_VERSION
        # 1. Check INI for explicit DELPHI_VERSION setting
        $DELPHI_VERSION = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "DELPHI_VERSION"
        # 2. Try to detect from source control path
        if ("$DELPHI_VERSION" -ceq "") {
            if ("$SOURCE_CONTROL_SOURCE_PATH" -like "*delphi 6*") { $DELPHI_VERSION = "6" }
            elseif ("$SOURCE_CONTROL_SOURCE_PATH" -like "*delphi 2007*") { $DELPHI_VERSION = "2007" }
            elseif ("$SOURCE_CONTROL_SOURCE_PATH" -like "*delphi xe2*") { $DELPHI_VERSION = "XE2" }
            elseif ("$SOURCE_CONTROL_SOURCE_PATH" -like "*delphi xe6*") { $DELPHI_VERSION = "XE6" }
            elseif ("$SOURCE_CONTROL_SOURCE_PATH" -like "*delphi 10.4*" -or "$SOURCE_CONTROL_SOURCE_PATH" -like "*sydney*") { $DELPHI_VERSION = "10.4" }
        }
        # 3. Default to 10.4 Sydney if not detected
        if ("$DELPHI_VERSION" -ceq "") {
            $DELPHI_VERSION = "10.4"
        }
        # 4. Always confirm with user — pre-select the detected/default version
        $delphiOptions = @("Delphi 6", "Delphi 2007", "Delphi XE2", "Delphi XE6", "Delphi 10.4 Sydney", "Cancel build")
        $delphiMap = @{ "6" = 0; "2007" = 1; "XE2" = 2; "XE6" = 3; "10.4" = 4 }
        $defaultIdx = $delphiMap[$DELPHI_VERSION] ?? 4
        Write-Log "Delphi auto-detected as '$DELPHI_VERSION' for '$INI_SECTION' — confirming with user..."
        $delphiChoice = Show-RadioMenu -Title "Confirm Delphi version for: $INI_SECTION" -Options $delphiOptions -DefaultIndex $defaultIdx
        switch ($delphiChoice) {
            0 { $DELPHI_VERSION = "6" }
            1 { $DELPHI_VERSION = "2007" }
            2 { $DELPHI_VERSION = "XE2" }
            3 { $DELPHI_VERSION = "XE6" }
            4 { $DELPHI_VERSION = "10.4" }
            default { throw "Build cancelled - no Delphi version selected" }
        }
        Write-Log "Using Delphi version: $DELPHI_VERSION"
        Write-Log "Delphi compiler version: $DELPHI_VERSION"
        if ("$SOURCE_CONTROL_SOURCE_PATH" -like "*delphi 6*") {
            $DELPHI_VERSION = "6"
        }
        if ("$SOURCE_CONTROL_SOURCE_PATH" -like "*delphi 2007*") {
            $DELPHI_VERSION = "2007"
        }
        if ("$SOURCE_CONTROL_SOURCE_PATH" -like "*delphi xe2*") {
            $DELPHI_VERSION = "XE2"
        }
        if ("$SOURCE_CONTROL_SOURCE_PATH" -like "*delphi xe6*") {
            $DELPHI_VERSION = "XE6"
        }
        if ("$DELPHI_VERSION" -ceq "") {
            $VAR_RESULT_TEXT = ""  # Initialise result variable
            foreach ($VAR_RESULT_TEXT in (Get-VaultFiles -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH" -Filter "Compile using *.txt")) {  # Find file describing compiler to be used
            }
            if ("$VAR_RESULT_TEXT" -like "*delphi 6*") {
                $DELPHI_VERSION = "6"
            }
            if ("$VAR_RESULT_TEXT" -like "*delphi 2007*") {
                $DELPHI_VERSION = "2007"
            }
            if ("$VAR_RESULT_TEXT" -like "*delphi xe2*") {
                $DELPHI_VERSION = "XE2"
            }
            if ("$VAR_RESULT_TEXT" -like "*delphi xe6*") {
                $DELPHI_VERSION = "XE6"
            }
            if ("$VAR_RESULT_TEXT" -like "*delphi 10.4*") {
                $DELPHI_VERSION = "10.4"
            }
        }
        if ("$DELPHI_VERSION" -ceq "") {
            throw "Unable to determine Delphi compiler version from the configuration file:  ${ABSOPENEDPROJECTDIR}PAApplications.ini"
        }

        #region Get the highest available project version from Source Control
        Write-Log "--- Get the highest available project version from Source Control ---"
        $SOURCE_CONTROL_LABEL = ""  # Set SOURCE_CONTROL_LABEL to a blank string
        $SOURCE_CONTROL_VERSION = ""  # Set SOURCE_CONTROL_VERSION to a blank string
        $SOURCE_CONTROL_LABEL = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "SOURCE_CONTROL_LABEL"  # SOURCE_CONTROL_LABEL
        $SOURCE_CONTROL_VERSION = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "SOURCE_CONTROL_VERSION"  # SOURCE_CONTROL_VERSION
        if ("$SOURCE_CONTROL_VERSION" -ceq "") {  # If SOURCE_CONTROL_VERSION not found in the ini file
            foreach ($VAR_RESULT_TEXT in (Get-VaultFiles -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH" -Filter "dcc32.cfg")) {  # Loop through all dcc32.cfg files (nearly all projects have one)
                $TEMP_VAR = "$VAR_RESULT_TEXT"  # Copy VAR_RESULT_TEXT into TEMP_VAR
                $TEMP_VAR = $TEMP_VAR.Replace('/', '\')  # Turn TEMP_VAR into file path
                $TEMP_VAR = Split-Path -Path "$TEMP_VAR" -Leaf  # Extract file name from TEMP_VAR
                $VAR_RESULT_TEXT = Get-SubstringBetween -Input "$VAR_RESULT_TEXT" -Start "$SOURCE_CONTROL_SOURCE_PATH" -End "Source/$TEMP_VAR"  # Remove SOURCE_CONTROL_SOURCE_PATH and TEMP_VAR from VAR_RESULT_TEXT
                $VAR_RESULT_TEXT = $VAR_RESULT_TEXT.Replace('/', ' ')  # Turn strip any / from VAR_RESULT_TEXT
                $VAR_RESULT_TEXT = ("$VAR_RESULT_TEXT").Trim()  # Trim VAR_RESULT_TEXT
                # Script block (JScript): Set SOURCE_CONTROL_VERSION to the higher of SOURCE_CONTROL_VERSION and the one just read
                if ($VAR_RESULT_TEXT -match '^\d+(?:\.\d+){0,3}$') {
                    [version]$newVersion = $VAR_RESULT_TEXT
                    [version]$oldVersion = '0.0.0.0'
                    if ($SOURCE_CONTROL_VERSION -match '^\d+(?:\.\d+){0,3}$') {
                        [version]$oldVersion = $SOURCE_CONTROL_VERSION
                    }
                    if ($newVersion -gt $oldVersion) {
                        $SOURCE_CONTROL_VERSION = $VAR_RESULT_TEXT
                    }
                }
                # [DISABLED] Log Message
            }
        }
        if ("$SOURCE_CONTROL_VERSION" -ceq "") {  # If SOURCE_CONTROL_VERSION not found in the dcc32.cfg file
            foreach ($VAR_RESULT_TEXT in (Get-VaultFiles -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH" -Filter "*.ini")) {  # Loop through all INI files (nearly all projects have an ini file)
                $TEMP_VAR = "$VAR_RESULT_TEXT"  # Copy VAR_RESULT_TEXT into TEMP_VAR
                $TEMP_VAR = $TEMP_VAR.Replace('/', '\')  # Turn TEMP_VAR into file path
                $TEMP_VAR = Split-Path -Path "$TEMP_VAR" -Leaf  # Extract file name from TEMP_VAR
                $VAR_RESULT_TEXT = Get-SubstringBetween -Input "$VAR_RESULT_TEXT" -Start "$SOURCE_CONTROL_SOURCE_PATH" -End "$TEMP_VAR"  # Remove SOURCE_CONTROL_SOURCE_PATH and TEMP_VAR from VAR_RESULT_TEXT
                $VAR_RESULT_TEXT = $VAR_RESULT_TEXT.Replace('/', ' ')  # Turn strip any / from VAR_RESULT_TEXT
                $VAR_RESULT_TEXT = ("$VAR_RESULT_TEXT").Trim()  # Trim VAR_RESULT_TEXT
                # Script block (JScript): Set SOURCE_CONTROL_VERSION to the higher of SOURCE_CONTROL_VERSION and the one just read
                if ($VAR_RESULT_TEXT -match '^\d+(?:\.\d+){0,3}$') {
                    [version]$newVersion = $VAR_RESULT_TEXT
                    [version]$oldVersion = '0.0.0.0'
                    if ($SOURCE_CONTROL_VERSION -match '^\d+(?:\.\d+){0,3}$') {
                        [version]$oldVersion = $SOURCE_CONTROL_VERSION
                    }
                    if ($newVersion -gt $oldVersion) {
                        $SOURCE_CONTROL_VERSION = $VAR_RESULT_TEXT
                    }
                }
                # [DISABLED] Log Message
            }
        }
        if ("$SOURCE_CONTROL_VERSION" -ne "") {  # If we have a SOURCE_CONTROL_VERSION
            $SOURCE_CONTROL_SOURCE_PATH = "$SOURCE_CONTROL_SOURCE_PATH/$SOURCE_CONTROL_VERSION"  # Append SOURCE_CONTROL_VERSION to SOURCE_CONTROL_SOURCE_PATH
        }
        #endregion Get the highest available project version from Source Control

        $BUILD_PATH = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "BUILD_PATH"  # BUILD_PATH
        if ("$BUILD_PATH" -ceq "") {
            throw "Unable to determine build path from the configuration file:  ${ABSOPENEDPROJECTDIR}PAApplications.ini"
        }
        if ("$BUILD_PATH" -notlike "C:\Builds\Under Development*") {
            throw "BUILD_PATH should start with C:\Builds\Under Development:  $BUILD_PATH"
        }
        $BUILD_TEMP_PATH = "$BUILD_LOCATION_PREFIX$SOURCE_CONTROL_SOURCE_PATH"  # Set BUILD_TEMP_PATH
        $BUILD_TEMP_PATH = ("$BUILD_TEMP_PATH").Replace('$/', '\')  # Replace $/ with \
        $BUILD_TEMP_PATH = ("$BUILD_TEMP_PATH").Replace('/', '\')  # Replace / with \
        # [DISABLED] If ... Then (If we have a SOURCE_CONTROL_VERSION) - If we have a SOURCE_CONTROL_VERSION
        $SUPPORTS_ORACLE = "Y"  # Set SUPPORTS_ORACLE
        $SUPPORTS_SUN4 = "Y"  # Set SUPPORTS_SUN4
        $SUPPORTS_SUN5 = "Y"  # Set SUPPORTS_SUN5
        $SUPPORTS_SUN6 = "Y"  # Set SUPPORTS_SUN6
        $SUPPORTS_ORACLE = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "SUPPORTS_ORACLE"  # SUPPORTS_ORACLE
        $SUPPORTS_SUN4 = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "SUPPORTS_SUN4"  # SUPPORTS_SUN4
        $SUPPORTS_SUN5 = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "SUPPORTS_SUN5"  # SUPPORTS_SUN5
        $SUPPORTS_SUN6 = Get-IniValue -Path "${ABSOPENEDPROJECTDIR}PAApplications.ini" -Section "$INI_SECTION" -Key "SUPPORTS_SUN6"  # SUPPORTS_SUN6
        $BUILD_INI = "$ABSOPENEDPROJECTDIR\Logs\$INI_SECTION.ini"  # Set BUILD_INI
        if (Test-Path "$ABSOPENEDPROJECTDIR\Logs") {  # 
        } else {
            New-Item -ItemType Directory -Path "$ABSOPENEDPROJECTDIR\Logs" -Force | Out-Null  # 
        }
        if (Test-Path "$BUILD_INI") {  # If BUILD_INI file exists
            $LAST_BUILD_DATE_TIME = Get-IniValue -Path "$BUILD_INI" -Section "$INI_SECTION" -Key "LAST_BUILD_DATE_TIME"  # LAST_BUILD_DATE_TIME
            if ("$LAST_BUILD_DATE_TIME" -eq "") {  # If LAST_BUILD_DATE_TIME not found
                $LAST_BUILD_DATE_TIME = "Script: Utilities.DateTimeToStr(Now)"  # Set LAST_BUILD_DATE_TIME
            }
        } else {
            Set-Content -Path "$BUILD_INI" -Value "[$INI_SECTION] LAST_BUILD_DATE_TIME= LAST_BUILD_VERSION="  # Create default file
            $LAST_BUILD_DATE_TIME = "Script: Utilities.DateTimeToStr(Now)"  # Set LAST_BUILD_DATE_TIME
        }
    } else {
        throw "Unable to locate configuration file:  ${ABSOPENEDPROJECTDIR}PAApplications.ini"
    }
    #endregion Read settings


    #region Build Standard setups (release and debug)
    Write-Log "--- Build Standard setups (release and debug) ---"
    if ("$SOURCE_CONTROL_LABEL" -ne "") {  # If SOURCE_CONTROL_LABEL is set
        $CONTINUE_LABEL_BUILD = Confirm-Action -Message "About to start a build of $PROJECT_TITLE by label. This build is really only suitable for rebuilding ActiveX setups, read below.  No unit tests and no PASQL tests will be run for this build due to version issues.  This is important! Any files picked up automatically by the Wise Install Script will be sourced from `"\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files`" and therefore most likely not match the original setup!  Due to Help.chm files not being versioned, the path to these will have to be entered set in PaApplications.ini file.  These values need to be defined in PAApplications.ini file. ; Base the build on sources matching this label SOURCE_CONTROL_LABEL=Bank Reconciliation - build 5.6.4.1 - 23/04/2014 3:44:21 PM ; Set to the highest number available in Source Control unless specified here SOURCE_CONTROL_VERSION=5.6 ; Overrides the standard help location HELP_FILE_LOCATION=C:\Builds\Released\Bank Reconciliation\5.6.4  These values have been found: SOURCE_CONTROL_LABEL = `"$SOURCE_CONTROL_LABEL`" SOURCE_CONTROL_VERSION = `"$SOURCE_CONTROL_VERSION`" HELP_FILE_LOCATION = `"$HELP_FILE_LOCATION`"  Everything will be build in this temporary folder will be `"$BUILD_TEMP_PATH`"  Continue build anyway? " -Default "True"  # 
        if ("$CONTINUE_LABEL_BUILD" -ne "TRUE") {
            return  # Stop Macro Execution - 
        }
    }

    #region Build Delphi projects
    Write-Log "--- Build Delphi projects ---"

    #region Get sources from Source Control
    Write-Log "--- Get sources from Source Control ---"
    if (Test-Path "$BUILD_TEMP_PATH") {  # If directory exists
        Invoke-DosCommand -Command "takeown /F $BUILD_TEMP_PATH /R"  # Take ownership of it
    }
    try {  # 
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH" -Recurse  # Delete old files if they exist
    } catch {  # 
        if (Get-Process -Name "Wise32.exe" -ErrorAction SilentlyContinue) {  # 
            Stop-Process -Name "Wise32.exe" -Force -ErrorAction SilentlyContinue  # 
            Remove-ItemSafe -Path "$BUILD_TEMP_PATH" -Recurse  # Delete old files if they exist
        }
    }
    New-Item -ItemType Directory -Path "$BUILD_TEMP_PATH" -Force | Out-Null  # Everything will be built here
    try {  # Exception may be raised if folder name has changed in Source Control
        if ("$SOURCE_CONTROL_LABEL" -ceq "") {  # If SOURCE_CONTROL_LABEL not set
            # [DISABLED] Try
            # [DISABLED] Catch
            Invoke-VaultCommand -Repository "SDG" -Command "SETWORKINGFOLDER" -Parameters "-forcesubfolderstoinherit  `"$SOURCE_CONTROL_SOURCE_PATH`" `"$BUILD_TEMP_PATH`""  # Set working folder
            Invoke-VaultGetLatest -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/*"  # Get latest project source
        } else {
            Invoke-VaultGetByLabel -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/*" -Label "$SOURCE_CONTROL_LABEL" -LocalPath "$BUILD_TEMP_PATH"  # Get project source by label
        }
    } catch {  # 
        throw "Unable to locate project in Source Control Error message: $TEMP_VAR"
    }
    #endregion Get sources from Source Control


    #region Get latest help file (if exists)
    Write-Log "--- Get latest help file (if exists) ---"
    if ("$SOURCE_CONTROL_LABEL" -ne "") {  # If SOURCE_CONTROL_LABEL is set
        if ("$HELP_FILE_LOCATION" -ne "") {  # If SOURCE_CONTROL_LABEL is set
            if (Test-Path "$HELP_FILE_LOCATION") {  # If directory exists
            } else {
                throw "Invalid path `"$HELP_FILE_LOCATION`""
            }
        }
    } else {
        if ("$INI_SECTION" -like "ePay Japanese*") {  # ePay Japanese quirk - has its own help file folder
            $HELP_FILE_LOCATION = "\\$VAULT_SERVER_ADDRESS\groups\QA\Online Help Files\$PROJECT_TITLE (Japanese)"
        } else {
            $HELP_FILE_LOCATION = "\\$VAULT_SERVER_ADDRESS\groups\QA\Online Help Files\$PROJECT_TITLE"
        }
    }
    if (Test-Path "$HELP_FILE_LOCATION") {  # If directory exists
        Copy-FileEx -Source "$HELP_FILE_LOCATION\*.chm" -Destination "$BUILD_TEMP_PATH" -Force  # Copy any .CHM file to the project folder
    }
    #endregion Get latest help file (if exists)


    #region Get latest dbExpress and Midas dlls
    Write-Log "--- Get latest dbExpress and Midas dlls ---"
    if ("$DELPHI_VERSION" -ceq "6") {  # If Delphi 6 project
        Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Setup Include Files\dbExpress\dbexp*da.dll" -Destination "$BUILD_TEMP_PATH\" -Force  # 
        Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Setup Include Files\winsys\midas.dll" -Destination "$BUILD_TEMP_PATH\" -Force  # 
    } else {
        if (("$DELPHI_VERSION" -eq "XE6") -and ("$DELPHI_VERSION" -eq "10.4")) {  # if XE6, 10.4
            if (Test-Path "$BUILD_TEMP_PATH\source\dcc32.cfg") {  # If dcc32.cfg file exists
                Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Setup Include Files\dbExpress\dbexp*da4?.dll" -Destination "$BUILD_TEMP_PATH\" -Force  # 
                Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Setup Include Files\winsys\midas.dll" -Destination "$BUILD_TEMP_PATH\" -Force  # 
            } else {
                Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Setup Include Files\dbExpress\win64\dbexp*da4?.dll" -Destination "$BUILD_TEMP_PATH\" -Force  # 
                Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Setup Include Files\winsys\win64\midas.dll" -Destination "$BUILD_TEMP_PATH\" -Force  # 
            }
            Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Setup Include Files\PA PGP Utils\PAPGPUtils*.dll" -Destination "$BUILD_TEMP_PATH\" -Force  # 
            # [DISABLED] Copy File(s)
        } else {  # XE2
            Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Setup Include Files\dbExpress\dbexp*da40.dll" -Destination "$BUILD_TEMP_PATH\" -Force  # 
            Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Setup Include Files\winsys\midas.dll" -Destination "$BUILD_TEMP_PATH\" -Force  # 
            Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Setup Include Files\PA PGP Utils\PAPGPUtils32.dll" -Destination "$BUILD_TEMP_PATH\" -Force  # 
        }
    }
    # [DISABLED] Copy File(s) (Copy dlls for PAUnit (XE2 project)) - Copy dlls for PAUnit (XE2 project)
    #endregion Get latest dbExpress and Midas dlls


    #region Update version numbers
    Write-Log "--- Update version numbers ---"

    #region Initialise variables
    Write-Log "--- Initialise variables ---"
    if ("$BUILD_VERSION_SET" -ne "Y") {  # If BUILD_VERSION_SET <> Y
        $SETUP_PASQL_FILE_NAME = ""  # Set SETUP_PASQL_FILE_NAME to ''
        $PASQL_SCRIPTS_HAVE_CHANGED = "FALSE"  # Set PASQL_SCRIPTS_HAVE_CHANGED to FALSE
        $V_MAJOR = "0"  # Set V_MAJOR to 0
        $V_MINOR = "0"  # Set V_MINOR to 0
        $V_RELEASE = "0"  # Set V_RELEASE to 0
        $V_BUILD = "0"  # Set V_BUILD to 0
        $BUILD_VERSION = "0.0.0.0"  # Set BUILD_VERSION to 0.0.0.0
        $IS_2_TIER = "TRUE"  # Set IS_2_TIER to TRUE
    }
    #endregion Initialise variables


    #region Get the highest version from project files and prompt for new version
    Write-Log "--- Get the highest version from project files and prompt for new version ---"
    if ("$BUILD_VERSION_SET" -ne "Y") {  # If BUILD_VERSION_SET <> Y
        if ("$DELPHI_VERSION" -ceq "10.4") {  # Delphi 10.4
            foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\source\*.dproj" -ErrorAction SilentlyContinue)) {  # Loop through DPROJ project files, project name in VAR_RESULT_TEXT
                $VAR_RESULT_TEXT = $__file.FullName
                # If XML Node/Attribute Exists: //*[local-name()='VerInfo_MajorVer']
                $FIND_RESULT = 'False'
                try {
                    [xml]$_xml = Get-Content -Path "$VAR_RESULT_TEXT" -Raw
                    $_nsMgr = [System.Xml.XmlNamespaceManager]::new($_xml.NameTable)
                    $_node = $_xml.SelectSingleNode("//*[local-name()='VerInfo_MajorVer']", $_nsMgr)
                    if ($_node) { $FIND_RESULT = 'True' }
                } catch { Write-Log "XML query failed: $_" }
                if ($FIND_RESULT -ceq 'True') {
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VerInfo_MajorVer'] "  # V_MAJOR
                } else {
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"MajorVer`"] "  # V_MAJOR
                }
                # If XML Node/Attribute Exists: //*[local-name()='VerInfo_MinorVer']
                $FIND_RESULT = 'False'
                try {
                    [xml]$_xml = Get-Content -Path "$VAR_RESULT_TEXT" -Raw
                    $_nsMgr = [System.Xml.XmlNamespaceManager]::new($_xml.NameTable)
                    $_node = $_xml.SelectSingleNode("//*[local-name()='VerInfo_MinorVer'] ", $_nsMgr)
                    if ($_node) { $FIND_RESULT = 'True' }
                } catch { Write-Log "XML query failed: $_" }
                if ($FIND_RESULT -ceq 'True') {
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VerInfo_MinorVer'] "  # V_MINOR
                } else {
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"MinorVer`"] "  # V_MINOR
                }
                # If XML Node/Attribute Exists: //*[local-name()='VerInfo_Release']
                $FIND_RESULT = 'False'
                try {
                    [xml]$_xml = Get-Content -Path "$VAR_RESULT_TEXT" -Raw
                    $_nsMgr = [System.Xml.XmlNamespaceManager]::new($_xml.NameTable)
                    $_node = $_xml.SelectSingleNode("//*[local-name()='VerInfo_Release']", $_nsMgr)
                    if ($_node) { $FIND_RESULT = 'True' }
                } catch { Write-Log "XML query failed: $_" }
                if ($FIND_RESULT -ceq 'True') {
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VerInfo_Release'] "  # V_RELEASE
                } else {
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"Release`"] "  # V_RELEASE
                }
                # If XML Node/Attribute Exists: //*[local-name()='VerInfo_Build']
                $FIND_RESULT = 'False'
                try {
                    [xml]$_xml = Get-Content -Path "$VAR_RESULT_TEXT" -Raw
                    $_nsMgr = [System.Xml.XmlNamespaceManager]::new($_xml.NameTable)
                    $_node = $_xml.SelectSingleNode("//*[local-name()='VerInfo_Build'] ", $_nsMgr)
                    if ($_node) { $FIND_RESULT = 'True' }
                } catch { Write-Log "XML query failed: $_" }
                if ($FIND_RESULT -ceq 'True') {
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VerInfo_Build'] "  # V_BUILD
                } else {
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"Build`"] "  # V_BUILD
                }
                # Script block (VBScript): Set BUILD_VERSION to higher of two versions
                $TEMP = "$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD"
                if ((Compare-Versions $TEMP $BUILD_VERSION) -gt 0) {
                  $BUILD_VERSION = $TEMP
                  Write-Log "BUILD_VERSION updated to: $BUILD_VERSION"
                }
            }
        } else {
            if ("$DELPHI_VERSION" -ceq "6") {  # Delphi 6
                foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\source\*.dof" -ErrorAction SilentlyContinue)) {  # Loop through DOF project files, project name in VAR_RESULT_TEXT
                    $VAR_RESULT_TEXT = $__file.FullName
                    $V_MAJOR = Get-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info" -Key "MajorVer"  # 
                    $V_MINOR = Get-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info" -Key "MinorVer"  # 
                    $V_RELEASE = Get-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info" -Key "Release"  # 
                    $V_BUILD = Get-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info" -Key "Build"  # 
                    # Script block (VBScript): Set SOURCE_CONTROL_VERSION to higher of two versions
                    $TEMP = "$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD"
                    if ((Compare-Versions $TEMP $SOURCE_CONTROL_VERSION) -gt 0) {
                      $SOURCE_CONTROL_VERSION = $TEMP
                      Write-Log "SOURCE_CONTROL_VERSION updated to: $SOURCE_CONTROL_VERSION"
                    }
                }
            } else {
                foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\source\*.dproj" -ErrorAction SilentlyContinue)) {  # Loop through DPROJ project files, project name in VAR_RESULT_TEXT
                    $VAR_RESULT_TEXT = $__file.FullName
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"MajorVer`"] "  # V_MAJOR
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"MinorVer`"] "  # V_MINOR
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"Release`"] "  # V_RELEASE
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"Build`"] "  # V_BUILD
                    # Script block (VBScript): Version comparison
                    $TEMP = "$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD"
                    if ((Compare-Versions $TEMP $SOURCE_CONTROL_VERSION) -gt 0) {
                      $SOURCE_CONTROL_VERSION = $TEMP
                      Write-Log "SOURCE_CONTROL_VERSION updated to: $SOURCE_CONTROL_VERSION"
                    }
                }
            }
        }
    }
    # To force specific version number, check both steps below
    # [DISABLED] Set/Reset Variable Value (Set BUILD_VERSION here to overcome issues with source files) - Set BUILD_VERSION here to overcome issues with source files
    # [DISABLED] Set/Reset Variable Value (Set BUILD_VERSION_SET = Y) - Set BUILD_VERSION_SET = Y
    # To force specific version number, check both steps above
    if ("$BUILD_VERSION_SET" -ne "Y") {  # If BUILD_VERSION_SET <> Y
        # Script block (VBScript): Save temporary numbers (incremented by 1)
        # Converted from DelphiScript: Save temporary numbers (incremented by 1)
        # Script block (VBScript): Parse version into V_MAJOR/MINOR/RELEASE/BUILD
        $versionParts = $BUILD_VERSION -split '\.'
        $V_MAJOR = [int]$versionParts[0]
        $V_MINOR = [int]$versionParts[1]
        $V_RELEASE = [int]$versionParts[2]
        $V_BUILD = [int]$versionParts[3]

        # Also create temp variables for incrementing
        $T_MAJ = $V_MAJOR + 1
        $T_MIN = $V_MINOR + 1
        $T_REL = $V_RELEASE + 1
        $T_BLD = $V_BUILD + 1

        Write-Log "Parsed version: $V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD"        if (("$NIGHTLY_BUILD" -ne "TRUE") -and ("$SOURCE_CONTROL_LABEL" -eq "")) {  # If not nightly build and SOURCE_CONTROL_LABEL is not set
            # TODO: String Substring - check parameters for $BUILD_YEAR  # Get the last 2 digits of current year
            # Radio Group: 
            $VAR_RESULT = Show-RadioMenu -Title "Set version for this build" -Options @(
                "$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD (no version change)",  # index=0
                "$V_MAJOR.$V_MINOR.$V_RELEASE.$T_BLD (standard build)",  # index=1
                "$V_MAJOR.$V_MINOR.$T_REL.0 (release build)",  # index=2
                "$V_MAJOR.$T_MIN.0.0 (minor build)",  # index=3
                "$BUILD_YEAR.1.0.0 (major build)",  # index=4
                "None of the above - cancel build process"  # index=5
            )
            if ("$VAR_RESULT" -ceq "5") {  # Cancel build
                # [DISABLED] Stop Macro Execution
                throw "Build cancelled by user"
            }
            if ("$VAR_RESULT" -ceq "1") {  # Standard build (build number incremented)
                # Script block (VBScript): V_BUILD + 1
                $V_BUILD = $V_BUILD + 1
                Write-Log "Version updated (standard build): $V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD"            }
            if ("$VAR_RESULT" -ceq "2") {  # Relase build (build number incremented)
                # Script block (VBScript): V_RELEASE + 1, V_BUILD = 0
                $V_RELEASE = $V_RELEASE + 1
                $V_BUILD = 0
                Write-Log "Version updated (release build): $V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD"                
                $PASQL_SCRIPTS_HAVE_CHANGED = "TRUE"  # Set PASQL_SCRIPTS_HAVE_CHANGED = True if changing "Build" number
            }
            if ("$VAR_RESULT" -ceq "3") {  # Minor build (minor number incremented)
                # Script block (VBScript): V_MINOR + 1, V_RELEASE = 0, V_BUILD = 0
                $V_MINOR = $V_MINOR + 1
                $V_RELEASE = 0
                $V_BUILD = 0
                Write-Log "Version updated (minor build): $V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD"
                $PASQL_SCRIPTS_HAVE_CHANGED = "TRUE"  # Set PASQL_SCRIPTS_HAVE_CHANGED = True if changing "Minor" number
            }
            if ("$VAR_RESULT" -ceq "4") {  # Major build (major number incremented)
                # Script block (VBScript): V_MAJOR = current year, V_MINOR = 1, V_RELEASE = 0, V_BUILD = 0
                $V_MAJOR = $CURRENT_YEAR
                $V_MINOR = 1
                $V_RELEASE = 0
                $V_BUILD = 0
                Write-Log "Version updated (major build): $V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD"
                $PASQL_SCRIPTS_HAVE_CHANGED = "TRUE"  # Set PASQL_SCRIPTS_HAVE_CHANGED = True if changing "Major" number
            }
            # [DISABLED] If ... Then (V_MAJOR.V_MINOR.V_RELEASE or V_MAJOR.V_MINOR compare to SOURCE_CONTROL_VERSION and SOURCE_CONTROL_LABEL is not set) - V_MAJOR.V_MINOR.V_RELEASE or V_MAJOR.V_MINOR compare to SOURCE_CONTROL_VERSION and SOURCE_CONTROL_LABEL is not set
        }
        $BUILD_VERSION = "$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD"  # Set BUILD_VERSION as per selected build type
    }
    $BUILD_TITLE = "Script: DateTimeToStr(Now)"  # Set BUILD_TITLE
    if (("$PROJECT_TITLE" -like "Advanced Inquiry*") -and ("$PROJECT_TITLE" -like "Archive Inquiry*")) {  # Building Advanced/Archive Inquiry
        $BUILD_TITLE = "Advanced/Archive Inquiry  $BUILD_VERSION - $BUILD_TITLE"  # Set BUILD_TITLE
        if ("$PROJECT_TITLE" -like "Advanced Inquiry*") {  # Set title only for the first project (Advanced Inquiry)
            $script:BuildTitle = "$BUILD_TITLE"  # Set build title
            Write-Log "Build Title: $BUILD_TITLE"
            $BUILD_VERSION_SET = "Y"  # Set BUILD_VERSION_SET = Y
        }
    } else {
        $BUILD_TITLE = "$PROJECT_TITLE  $BUILD_VERSION - $BUILD_TITLE"  # Set BUILD_TITLE
        $script:BuildTitle = "$BUILD_TITLE"  # Set build title
        Write-Log "Build Title: $BUILD_TITLE"
    }
    $RELEASE_VERSION = "$V_MAJOR.$V_MINOR.$V_RELEASE"  # Set RELEASE_VERSION as per selected build type
    $BUILD_PATH = "$BUILD_PATH\$RELEASE_VERSION"  # Append RELEASE_VERSION to the BUILD_PATH
    #endregion Get the highest version from project files and prompt for new version

    if ("$NIGHTLY_BUILD" -ne "TRUE") {  # If not nightly build

        #region Check if any PASQL script has changed
        Write-Log "--- Check if any PASQL script has changed ---"
        $VAR_RESULT = 0
        foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\scripts\*setup.pasql" -ErrorAction SilentlyContinue)) {  # Check for presence of *setup.pasql file
            $SETUP_PASQL_FILE_NAME = $__file.FullName
            $VAR_RESULT++
        }
        if ("$VAR_RESULT" -ne "") {  # If more than 1 *setup.pasql found
            throw "Only one *Setup.pasq file should exist. $VAR_RESULT found  BuildStudio script needs fixing :)"
        }
        if (("$PASQL_SCRIPTS_HAVE_CHANGED" -ne "TRUE") -and ("$VAR_RESULT" -ne "0")) {  # If not PASQL_SCRIPTS_HAVE_CHANGED yet set
            $SETUP_PASQL_DATE = (Get-Item "$SETUP_PASQL_FILE_NAME").LastWriteTime.ToString()  # Set SETUP_PASQL_DATE from setup.pasql file
            # Script block (DelphiScript): Compare setup and last-build dates
            $setupDate = [datetime]::MinValue
            $buildDate = [datetime]::MinValue
            if ([datetime]::TryParse("$SETUP_PASQL_DATE", [ref]$setupDate) -and [datetime]::TryParse("$LAST_BUILD_DATE_TIME", [ref]$buildDate)) {
                if ($setupDate -ge $buildDate) {
                    $PASQL_SCRIPTS_HAVE_CHANGED = "TRUE"
                }
            }
            if (("$PASQL_SCRIPTS_HAVE_CHANGED" -ne "TRUE") -and ("$VAR_RESULT" -ne "0")) {  # If not PASQL_SCRIPTS_HAVE_CHANGED yet set
                $VAR_RESULT = 0
                foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\scripts\*.pasql" -ErrorAction SilentlyContinue)) {  # Compare modification date of each PASQL file with SETUP_PASQ_DATE
                    $VAR_RESULT_TEXT = $__file.FullName
                    $VAR_RESULT++
                    if ("$VAR_RESULT_TEXT" -ne "$SETUP_PASQL_FILE_NAME") {  # If a more recent file found set PASQL_SCRIPTS_HAVE_CHANGED to TRUE
                        $TEMP_VAR = (Get-Item "$VAR_RESULT_TEXT").LastWriteTime.ToString()  # Get last modified date of each PASQL
                        # Script block (DelphiScript): Compare PASQL file date and setup date
                        $setupDate = [datetime]::MinValue
                        $fileDate = [datetime]::MinValue
                        if ([datetime]::TryParse("$SETUP_PASQL_DATE", [ref]$setupDate) -and [datetime]::TryParse("$TEMP_VAR", [ref]$fileDate)) {
                            if ($fileDate -ge $setupDate) {
                                $PASQL_SCRIPTS_HAVE_CHANGED = "TRUE"
                            }
                        }
                        if ("$PASQL_SCRIPTS_HAVE_CHANGED" -eq "TRUE") {
                            break
                        }
                    }
                }
            }
        }
        #endregion Check if any PASQL script has changed


        #region Check if the applications is 3 tier
        Write-Log "--- Check if the applications is 3 tier ---"
        if (Test-Path "$BUILD_TEMP_PATH\*server.ini") {  # 
            $IS_2_TIER = "FALSE"  # Set IS_2_TIER to FALSE
        }
        #endregion Check if the applications is 3 tier


        #region Update version numbers
        Write-Log "--- Update version numbers ---"
        if ("$SOURCE_CONTROL_LABEL" -ceq "") {  # If SOURCE_CONTROL_LABEL is not set

            #region Update setup.pasql file if needed
            Write-Log "--- Update setup.pasql file if needed ---"
            if ("$PASQL_SCRIPTS_HAVE_CHANGED" -eq "TRUE") {  # If PASQL scripts have changed
                if (Test-Path "$SETUP_PASQL_FILE_NAME") {  # If PASQL script exists
                    Copy-FileEx -Source "$SETUP_PASQL_FILE_NAME" -Destination "$SETUP_PASQL_FILE_NAME.test" -Force  # Make a copy of SETUP_PASQL_FILE_NAME to use in PASQL tests
                    $TEMP_VAR = Split-Path -Path "$SETUP_PASQL_FILE_NAME" -Leaf  # Set TEMP_VAR to the file part of SETUP_PASQL_FILE_NAME
                    Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/scripts/$TEMP_VAR" -Host "$VAULT_SERVER_ADDRESS"  # 
                    # Find/replace split into separate calls in case not all replacements are required
                    Replace-InFile -Path "$SETUP_PASQL_FILE_NAME" -Find "^AppDBVersion.*= `"\d+\.\d+\.\d+\.\d+`"" -Replace "AppDBVersion         = `"$BUILD_VERSION`""  # Set AppDBVersion
                    Replace-InFile -Path "$SETUP_PASQL_FILE_NAME" -Find "^AppMinServerVersion.*= `"\d+\.\d+\.\d+\.\d+`"" -Replace "AppMinServerVersion  = `"$BUILD_VERSION`""  # Set AppMinServerVersion
                    if ("$IS_2_TIER" -ne "TRUE") {
                        Replace-InFile -Path "$SETUP_PASQL_FILE_NAME" -Find "^AppMinClientVersion.*= `"\d+\.\d+\.\d+\.\d+`"" -Replace "AppMinClientVersion  = `"$BUILD_VERSION`""  # Set AppMinClientVersion
                    }
                }
            }
            #endregion Update setup.pasql file if needed


            #region Update project files
            Write-Log "--- Update project files ---"
            # Script block (DelphiScript): Build project icon name
            $ICON_FILE_NAME = ($PROJECT_TITLE -replace '\s+', '') + ".ico"
            if (Test-Path "$BUILD_TEMP_PATH\source\$ICON_FILE_NAME") {  # Check if source\ICON_FILE_NAME file exists
            } else {
                Invoke-VaultGetLatest -Repository "SDG" -Path "$/Framework/Icons/APAICON.ico" -LocalFolder "$BUILD_TEMP_PATH\source"  # Get APAICON.ico (for building RES files)
                $ICON_FILE_NAME = "APAICON.ICO"  # Set ICON_FILE_NAME to APAICON.ICO
            }
            if ("$DELPHI_VERSION" -ceq "6") {  # Delphi 6
                foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\source\*.dof" -ErrorAction SilentlyContinue)) {  # Loop through DOF project files, project name in VAR_RESULT_TEXT
                    $VAR_RESULT_TEXT = $__file.FullName
                    $TEMP_VAR = Split-Path -Path "$VAR_RESULT_TEXT" -Leaf  # Set TEMP_VAR to the file part of file
                    Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/$TEMP_VAR" -Host "$VAULT_SERVER_ADDRESS"  # 
                    Set-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info" -Key "MajorVer" -Value "$V_MAJOR"  # Set MajorVer
                    Set-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info" -Key "MinorVer" -Value "$V_MINOR"  # Set MinorVer
                    Set-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info" -Key "Release" -Value "$V_RELEASE"  # Set Release
                    Set-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info" -Key "Build" -Value "$V_BUILD"  # Set Build
                    Set-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info Keys" -Key "FileVersion" -Value "$BUILD_VERSION"  # Set FileVersion
                    Set-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info Keys" -Key "ProductVersion" -Value "$RELEASE_VERSION"  # Set ProductVersion
                    Set-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info Keys" -Key "LegalCopyright" -Value "Copyright © 1991 - $CURRENT_YEAR"  # Set LegalCopyright
                    $FILE_DESCRIPTION = Get-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info Keys" -Key "FileDescription"  # Set FILE_DESCRIPTION
                    $INTERNAL_NAME = Get-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info Keys" -Key "InternalName"  # Set INTERNAL_NAME
                    $PRODUCT_NAME = Get-IniValue -Path "$VAR_RESULT_TEXT" -Section "Version Info Keys" -Key "ProductName"  # Set PRODUCT_NAME
                    $VAR_RESULT_TEXT = [System.IO.Path]::ChangeExtension("$VAR_RESULT_TEXT", '.res')  # Change file name, replace DOF extension with RES
                    if (Test-Path "$VAR_RESULT_TEXT") {  # If a RES file for this project exists
                        $TEMP_VAR = Split-Path -Path "$VAR_RESULT_TEXT" -Leaf  # Set TEMP_VAR to the file part of file
                        Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/$TEMP_VAR" -Host "$VAULT_SERVER_ADDRESS"  # 
                        $VAR_RESULT_TEXT = [System.IO.Path]::ChangeExtension("$VAR_RESULT_TEXT", '.rc')  # Change file name, replace DOF extension with RC
                        Set-Content -Path "$VAR_RESULT_TEXT" -Value "MAINICON ICON `"$ICON_FILE_NAME`" APAICON ICON `"$ICON_FILE_NAME`"  1 VERSIONINFO  FILEVERSION $V_MAJOR,$V_MINOR,$V_RELEASE,$V_BUILD  PRODUCTVERSION $V_MAJOR,$V_MINOR,$V_RELEASE  FILEFLAGSMASK 0x3fL  FILEFLAGS 0x0L  FILEOS 0x4L  FILETYPE 0x1L  FILESUBTYPE 0x0L BEGIN     BLOCK `"StringFileInfo`"     BEGIN         BLOCK `"0c0904e4`"         BEGIN             VALUE `"CompanyName`", `"Professional Advantage Pty. Ltd.`"             VALUE `"FileDescription`", `"$FILE_DESCRIPTION`"             VALUE `"FileVersion`", `"$BUILD_VERSION`"             VALUE `"LegalCopyright`", `"©1991 - $CURRENT_YEAR`"             VALUE `"InternalName`", `"$INTERNAL_NAME`"             VALUE `"ProductName`", `"$PRODUCT_NAME`"             VALUE `"ProductVersion`", `"$RELEASE_VERSION`"         END     END     BLOCK `"VarFileInfo`"     BEGIN         VALUE `"Translation`", 0xc09, 1252     END END"  # Create RC file
                        Invoke-DosCommand -Command "`"C:\Program Files\Microsoft SDKs\Windows\v7.0A\bin\rc.exe`" /v /r /l0c09 /c1252 `"$VAR_RESULT_TEXT`"" -WorkingDirectory "$BUILD_TEMP_PATH\source"  # Build the RES file
                    }
                }
            } else {
                foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\source\*.dproj" -ErrorAction SilentlyContinue)) {  # Loop through DPROJ project files, project name in VAR_RESULT_TEXT
                    $VAR_RESULT_TEXT = $__file.FullName
                    $TEMP_VAR = Split-Path -Path "$VAR_RESULT_TEXT" -Leaf  # Set TEMP_VAR to the file part of file
                    Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/$TEMP_VAR" -Host "$VAULT_SERVER_ADDRESS"  # 
                    if ("$DELPHI_VERSION" -ceq "10.4") {  # Delphi 10.4
                        # If XML Node/Attribute Exists: //*[local-name()='VerInfo_MajorVer']
                        $FIND_RESULT = 'False'
                        try {
                            [xml]$_xml = Get-Content -Path "$VAR_RESULT_TEXT" -Raw
                            $_nsMgr = [System.Xml.XmlNamespaceManager]::new($_xml.NameTable)
                            $_node = $_xml.SelectSingleNode("//*[local-name()='VerInfo_MajorVer']", $_nsMgr)
                            if ($_node) { $FIND_RESULT = 'True' }
                        } catch { Write-Log "XML query failed: $_" }
                        if ($FIND_RESULT -ceq 'True') {
                            Set-XmlValue -Path "" -XPath "//*[local-name()='VerInfo_MajorVer']" -Value "$V_MAJOR"  # Set MajorVer
                        } else {
                            Set-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"MajorVer`"]" -Value "$V_MAJOR"  # Set MajorVer
                        }
                        # If XML Node/Attribute Exists: //*[local-name()='VerInfo_MinorVer']
                        $FIND_RESULT = 'False'
                        try {
                            [xml]$_xml = Get-Content -Path "$VAR_RESULT_TEXT" -Raw
                            $_nsMgr = [System.Xml.XmlNamespaceManager]::new($_xml.NameTable)
                            $_node = $_xml.SelectSingleNode("//*[local-name()='VerInfo_MinorVer'] ", $_nsMgr)
                            if ($_node) { $FIND_RESULT = 'True' }
                        } catch { Write-Log "XML query failed: $_" }
                        if ($FIND_RESULT -ceq 'True') {
                            Set-XmlValue -Path "" -XPath "//*[local-name()='VerInfo_MinorVer']" -Value "$V_MINOR"  # Set MinorVer
                        } else {
                            Set-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"MinorVer`"]" -Value "$V_MINOR"  # Set MinorVer
                        }
                        # If XML Node/Attribute Exists: //*[local-name()='VerInfo_Release']
                        $FIND_RESULT = 'False'
                        try {
                            [xml]$_xml = Get-Content -Path "$VAR_RESULT_TEXT" -Raw
                            $_nsMgr = [System.Xml.XmlNamespaceManager]::new($_xml.NameTable)
                            $_node = $_xml.SelectSingleNode("//*[local-name()='VerInfo_Release']", $_nsMgr)
                            if ($_node) { $FIND_RESULT = 'True' }
                        } catch { Write-Log "XML query failed: $_" }
                        if ($FIND_RESULT -ceq 'True') {
                            Set-XmlValue -Path "" -XPath "//*[local-name()='VerInfo_Release']" -Value "$V_RELEASE"  # Set Release
                        } else {
                            Set-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"Release`"]" -Value "$V_RELEASE"  # Set Release
                        }
                        # If XML Node/Attribute Exists: //*[local-name()='VerInfo_Build']
                        $FIND_RESULT = 'False'
                        try {
                            [xml]$_xml = Get-Content -Path "$VAR_RESULT_TEXT" -Raw
                            $_nsMgr = [System.Xml.XmlNamespaceManager]::new($_xml.NameTable)
                            $_node = $_xml.SelectSingleNode("//*[local-name()='VerInfo_Build'] ", $_nsMgr)
                            if ($_node) { $FIND_RESULT = 'True' }
                        } catch { Write-Log "XML query failed: $_" }
                        if ($FIND_RESULT -ceq 'True') {
                            Set-XmlValue -Path "" -XPath "//*[local-name()='VerInfo_Build']" -Value "$V_BUILD"  # Set Build
                        } else {
                            Set-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"Build`"]" -Value "$V_BUILD"  # Set Build
                        }
                    } else {
                        Set-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"MajorVer`"]" -Value "$V_MAJOR"  # Set MajorVer
                        Set-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"MinorVer`"]" -Value "$V_MINOR"  # Set MinorVer
                        Set-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"Release`"]" -Value "$V_RELEASE"  # Set Release
                        Set-XmlValue -Path "" -XPath "//*[local-name()='VersionInfo'][@Name=`"Build`"]" -Value "$V_BUILD"  # Set Build
                        Set-XmlValue -Path "" -XPath "//*[local-name()='VersionInfoKeys'][@Name=`"FileVersion`"]" -Value "$BUILD_VERSION"  # Set FileVersion
                        Set-XmlValue -Path "" -XPath "//*[local-name()='VersionInfoKeys'][@Name=`"ProductVersion`"]" -Value "$RELEASE_VERSION"  # Set ProductVersion
                        Set-XmlValue -Path "" -XPath "//*[local-name()='VersionInfoKeys'][@Name=`"LegalCopyright`"]" -Value "©1991 - $CURRENT_YEAR"  # Set LegalCopyright
                    }
                    # [DISABLED] Find/Replace in File (Set version numbers, copyright etc) - Set version numbers, copyright etc
                    # [DISABLED] Find/Replace in File (Remove damaged EurekaLog section) - Remove damaged EurekaLog section
                    # [DISABLED] Write to File (Append EUREKA_LOG_CONFIG_SECTION) - Append EUREKA_LOG_CONFIG_SECTION
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VersionInfoKeys'][@Name=`"FileDescription`"]"  # Set FILE_DESCRIPTION
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VersionInfoKeys'][@Name=`"InternalName`"]"  # Set INTERNAL_NAME
                    $None = Get-XmlValue -Path "" -XPath "//*[local-name()='VersionInfoKeys'][@Name=`"ProductName`"]"  # Set PRODUCT_NAME
                    if ((("$DELPHI_VERSION" -eq "XE2") -and ("$DELPHI_VERSION" -eq "XE6")) -and ("$DELPHI_VERSION" -eq "10.4")) {  # Delphi XE2, XE6, 10.4
                        if (("$VAR_RESULT_TEXT" -notlike "*ISAPI*") -and ("$VAR_RESULT_TEXT" -notlike "*REPORTQUERYBUILDER*")) {  # If project is not a  dll
                            # If XML Node/Attribute Exists: 
                            $NOT_A_DLL = 'False'
                            try {
                                [xml]$_xml = Get-Content -Path "$VAR_RESULT_TEXT" -Raw
                                $_nsMgr = [System.Xml.XmlNamespaceManager]::new($_xml.NameTable)
                                $_node = $_xml.SelectSingleNode("//*[local-name()='Icon_MainIcon']", $_nsMgr)
                                if ($_node) { $NOT_A_DLL = 'True' }
                            } catch { Write-Log "XML query failed: $_" }
                            if ($NOT_A_DLL -ceq 'True') {
                                Set-XmlValue -Path "" -XPath "//*[local-name()='Icon_MainIcon']" -Value "$ICON_FILE_NAME"  # Set Icon
                            }
                        }
                        Set-XmlValue -Path "" -XPath "//*[local-name()='VerInfo_Keys']" -Value "CompanyName=Professional Advantage Pty. Ltd.;FileDescription=$FILE_DESCRIPTION;FileVersion=$BUILD_VERSION;InternalName=$INTERNAL_NAME;LegalCopyright=©1991 - $CURRENT_YEAR;LegalTrademarks=;OriginalFilename=;ProductName=$PRODUCT_NAME;ProductVersion=$RELEASE_VERSION;Comments="  # Set Version info keys
                    }
                    $VAR_RESULT_TEXT = [System.IO.Path]::ChangeExtension("$VAR_RESULT_TEXT", '.res')  # Change file name, replace DPROJ extension with RES
                    if (Test-Path "$VAR_RESULT_TEXT") {  # If a RES file for this project exists
                        $TEMP_VAR = Split-Path -Path "$VAR_RESULT_TEXT" -Leaf  # Set TEMP_VAR to the file part of file
                        Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/$TEMP_VAR" -Host "$VAULT_SERVER_ADDRESS"  # 
                        $VAR_RESULT_TEXT = [System.IO.Path]::ChangeExtension("$VAR_RESULT_TEXT", '.rc')  # Change file name, replace DPROJ extension with RC
                        if (("$DELPHI_VERSION" -eq "XE2") -and ("$DELPHI_VERSION" -eq "XE6")) {  # XE2
                            Set-Content -Path "$VAR_RESULT_TEXT" -Value "MAINICON ICON `"$ICON_FILE_NAME`"  1 VERSIONINFO  FILEVERSION $V_MAJOR,$V_MINOR,$V_RELEASE,$V_BUILD  PRODUCTVERSION $V_MAJOR,$V_MINOR,$V_RELEASE  FILEFLAGSMASK 0x3fL  FILEFLAGS 0x0L  FILEOS 0x4L  FILETYPE 0x1L  FILESUBTYPE 0x0L BEGIN     BLOCK `"StringFileInfo`"     BEGIN         BLOCK `"0c0904e4`"         BEGIN             VALUE `"CompanyName`", `"Professional Advantage Pty. Ltd.`"             VALUE `"FileDescription`", `"$FILE_DESCRIPTION`"             VALUE `"FileVersion`", `"$BUILD_VERSION`"             VALUE `"LegalCopyright`", `"©1991 - $CURRENT_YEAR`"             VALUE `"InternalName`", `"$INTERNAL_NAME`"             VALUE `"ProductName`", `"$PRODUCT_NAME`"             VALUE `"ProductVersion`", `"$RELEASE_VERSION`"         END     END     BLOCK `"VarFileInfo`"     BEGIN         VALUE `"Translation`", 0xc09, 1252     END END"  # Create RC file
                        } else {
                            Set-Content -Path "$VAR_RESULT_TEXT" -Value "MAINICON ICON `"$ICON_FILE_NAME`"  1 VERSIONINFO  FILEVERSION $V_MAJOR,$V_MINOR,$V_RELEASE,$V_BUILD  PRODUCTVERSION $V_MAJOR,$V_MINOR,$V_RELEASE  FILEFLAGSMASK 0x3fL  FILEFLAGS 0x0L  FILEOS 0x4L  FILETYPE 0x1L  FILESUBTYPE 0x0L BEGIN     BLOCK `"StringFileInfo`"     BEGIN         BLOCK `"0c0904e4`"         BEGIN             VALUE `"CompanyName`", `"Professional Advantage Pty. Ltd.`"             VALUE `"FileDescription`", `"$FILE_DESCRIPTION`"             VALUE `"FileVersion`", `"$BUILD_VERSION`"             VALUE `"LegalCopyright`", `"©1991 - $CURRENT_YEAR`"             VALUE `"InternalName`", `"$INTERNAL_NAME`"             VALUE `"ProductName`", `"$PRODUCT_NAME`"             VALUE `"ProductVersion`", `"$RELEASE_VERSION`"         END     END     BLOCK `"VarFileInfo`"     BEGIN         VALUE `"Translation`", 0xc09, 1252     END END"  # Create RC file
                        }
                        Invoke-DosCommand -Command "`"C:\Program Files (x86)\Windows Kits\8.1\bin\x64\rc.exe`" /v /r /l0c09 /c1252 `"$VAR_RESULT_TEXT`"" -WorkingDirectory "$BUILD_TEMP_PATH\source"  # Build the RES file
                    }
                }
            }
            #endregion Update project files


            #region Update shared versions file
            Write-Log "--- Update shared versions file ---"
            if ("$DELPHI_VERSION" -ceq "6") {  # If Delphi 6
                if ("$IS_2_TIER" -ne "TRUE") {  # 3 tier
                    if (Test-Path "$BUILD_TEMP_PATH\source\shared\Version.pas") {  # If source\shared\Version.pas file exists
                        Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/shared//Version.pas" -Host "$VAULT_SERVER_ADDRESS"  # 
                        # Find/replace split into separate calls in case not all replacements are required
                        Replace-InFile -Path "$BUILD_TEMP_PATH\source\shared\Version.pas" -Find "MIN_SERVER_VERSION .*= '\d+\.\d+\.\d+\.\d+';" -Replace "MIN_SERVER_VERSION          = '$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD';"  # Set MIN_SERVER_VERSION
                        Replace-InFile -Path "$BUILD_TEMP_PATH\source\shared\Version.pas" -Find "MIN_CLIENT_VERSION .*= '\d+\.\d+\.\d+\.\d+';" -Replace "MIN_CLIENT_VERSION          = '$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD';"  # Set MIN_CLIENT_VERSION
                        if ("$PASQL_SCRIPTS_HAVE_CHANGED" -eq "TRUE") {  # If PASQL scripts have changed
                            Replace-InFile -Path "$BUILD_TEMP_PATH\source\shared\Version.pas" -Find "MIN_DATABASE_VERSION .*= '\d+\.\d+\.\d+\.\d+';" -Replace "MIN_DATABASE_VERSION        = '$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD';"  # Set MIN_DATABASE_VERSION
                        }
                    }
                    if (Test-Path "$BUILD_TEMP_PATH\source\shared\Application\Version.pas") {  # If source\shared\Application\Version.pas file exists
                        Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/shared/application/Version.pas" -Host "$VAULT_SERVER_ADDRESS"  # 
                        # Find/replace split into separate calls in case not all replacements are required
                        Replace-InFile -Path "$BUILD_TEMP_PATH\source\shared\Application\Version.pas" -Find "MIN_SERVER_VERSION .*= '\d+\.\d+\.\d+\.\d+';" -Replace "MIN_SERVER_VERSION          = '$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD';"  # Set MIN_SERVER_VERSION
                        Replace-InFile -Path "$BUILD_TEMP_PATH\source\shared\Application\Version.pas" -Find "MIN_CLIENT_VERSION .*= '\d+\.\d+\.\d+\.\d+';" -Replace "MIN_CLIENT_VERSION          = '$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD';"  # Set MIN_CLIENT_VERSION
                        if ("$PASQL_SCRIPTS_HAVE_CHANGED" -eq "TRUE") {  # If PASQL scripts have changed
                            Replace-InFile -Path "$BUILD_TEMP_PATH\source\shared\Application\Version.pas" -Find "MIN_DATABASE_VERSION .*= '\d+\.\d+\.\d+\.\d+';" -Replace "MIN_DATABASE_VERSION        = '$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD';"  # Set MIN_DATABASE_VERSION
                        }
                    }
                }
            } else {
                if ("$IS_2_TIER" -eq "TRUE") {  # 2 tier
                    if (Test-Path "$BUILD_TEMP_PATH\source\application\AppVersionConsts.pas") {  # If source\application\AppVersionConsts.pas file exists
                        Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/application/AppVersionConsts.pas" -Host "$VAULT_SERVER_ADDRESS"  # 
                        if ("$PASQL_SCRIPTS_HAVE_CHANGED" -eq "TRUE") {  # If PASQL scripts have changed
                            Replace-InFile -Path "$BUILD_TEMP_PATH\source\application\AppVersionConsts.pas" -Find "MIN_DATABASE_VERSION .*= '\d+\.\d+\.\d+\.\d+';" -Replace "MIN_DATABASE_VERSION        = '$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD';"  # Set MIN_DATABASE_VERSION
                        }
                    }
                } else {
                    if (Test-Path "$BUILD_TEMP_PATH\source\shared\application\AppVersionConsts.pas") {  # If source\shared\application\AppVersionConsts.pas file exists
                        Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/shared/application/AppVersionConsts.pas" -Host "$VAULT_SERVER_ADDRESS"  # 
                        # Find/replace split into separate calls in case not all replacements are required
                        Replace-InFile -Path "$BUILD_TEMP_PATH\source\shared\application\AppVersionConsts.pas" -Find "MIN_SERVER_VERSION *= '\d+\.\d+\.\d+\.\d+';" -Replace "MIN_SERVER_VERSION          = '$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD';"  # Set MIN_SERVER_VERSION
                        Replace-InFile -Path "$BUILD_TEMP_PATH\source\shared\application\AppVersionConsts.pas" -Find "MIN_SERVER_VERSION_FOR_AUTO *= '\d+\.\d+\.\d+\.\d+';" -Replace "MIN_SERVER_VERSION_FOR_AUTO = '$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD';"  # Set MIN_SERVER_VERSION_FOR_AUTO (Pillar SunSystems Interface)
                        Replace-InFile -Path "$BUILD_TEMP_PATH\source\shared\application\AppVersionConsts.pas" -Find "MIN_CLIENT_VERSION .*= '\d+\.\d+\.\d+\.\d+';" -Replace "MIN_CLIENT_VERSION          = '$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD';"  # Set MIN_CLIENT_VERSION
                        if ("$PASQL_SCRIPTS_HAVE_CHANGED" -eq "TRUE") {  # If PASQL scripts have changed
                            Replace-InFile -Path "$BUILD_TEMP_PATH\source\shared\application\AppVersionConsts.pas" -Find "MIN_DATABASE_VERSION .*= '\d+\.\d+\.\d+\.\d+';" -Replace "MIN_DATABASE_VERSION        = '$V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD';"  # Set MIN_DATABASE_VERSION
                        }
                    }
                }
            }
            #endregion Update shared versions file


            #region Update BuildSetupConsts.vbs if needed (TODO: this file is not used by Build Studio!)
            Write-Log "--- Update BuildSetupConsts.vbs if needed (TODO: this file is not used by Build Studio!) ---"
            if (Test-Path "$BUILD_TEMP_PATH\setup\BuildSetupConsts.vbs") {  # If BuildSetupConsts.vbs file exists
                Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/setup/BuildSetupConsts.vbs" -Host "$VAULT_SERVER_ADDRESS"  # 
                # Find/replace split into separate calls in case not all replacements are required
                Replace-InFile -Path "$BUILD_TEMP_PATH\setup\BuildSetupConsts.vbs" -Find "DistributeTo.*= `".*`"$" -Replace "DistributeTo = `"\\\\$VAULT_SERVER_ADDRESS\\ForQA\\$PROJECT_TITLE\\$RELEASE_VERSION`""  # Set DistributeTo
                if ("$PROJECT_TITLE" -like "Contract and Service Billing*") {  # CSB quirk - release folder does not match project title (PROJECT_TITLE)
                    Replace-InFile -Path "$BUILD_TEMP_PATH\setup\BuildSetupConsts.vbs" -Find "DistributeTo.*= `".*`"$" -Replace "DistributeTo = `"\\\\$VAULT_SERVER_ADDRESS\\ForQA\\CSB\\$RELEASE_VERSION`""  # Set DistributeTo
                }
            }
            #endregion Update BuildSetupConsts.vbs if needed (TODO: this file is not used by Build Studio!)

        }
        #endregion Update version numbers

        if ("$SOURCE_CONTROL_LABEL" -ceq "") {  # If SOURCE_CONTROL_LABEL is not set

            #region Update About dialog's copyright year if needed
            Write-Log "--- Update About dialog's copyright year if needed ---"
            if ("$DELPHI_VERSION" -ceq "6") {  # Delphi 6
                if ("$IS_2_TIER" -eq "TRUE") {  # 2 tier
                    if (Test-Path "$BUILD_TEMP_PATH\Source\pa shared\paAboutDialog.pas") {  # If source\pa shared\paAboutDialog.pas file exists
                        $VAR_RESULT = Find-InFile -Path "$BUILD_TEMP_PATH\Source\PA Shared\paAboutDialog.pas" -Find "PA_COPYRIGHT_YEAR.*= *'$CURRENT_YEAR';"  # Set PA_COPYRIGHT_YEAR
                        if ("$VAR_RESULT" -le "1") {  # Only modify this file if the current year does not match
                            Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/pa shared/paAboutDialog.pas" -Host "$VAULT_SERVER_ADDRESS"  # 
                            Replace-InFile -Path "$BUILD_TEMP_PATH\Source\PA Shared\paAboutDialog.pas" -Find "PA_COPYRIGHT_YEAR.*= *'\d\d\d\d';" -Replace "PA_COPYRIGHT_YEAR = '$CURRENT_YEAR';"  # Set PA_COPYRIGHT_YEAR
                        }
                    }
                } else {  # 3 tier
                    if (Test-Path "$BUILD_TEMP_PATH\Source\Client\pa shared\paAboutDialog.pas") {  # If source\Client\pa shared\paAboutDialog.pas file exists
                        $VAR_RESULT = Find-InFile -Path "$BUILD_TEMP_PATH\Source\Client\PA Shared\paAboutDialog.pas" -Find "PA_COPYRIGHT_YEAR.*= *'$CURRENT_YEAR';"  # Set PA_COPYRIGHT_YEAR
                        if ("$VAR_RESULT" -le "1") {  # Only modify this file if the current year does not match
                            Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/client/pa shared/paAboutDialog.pas" -Host "$VAULT_SERVER_ADDRESS"  # 
                            Replace-InFile -Path "$BUILD_TEMP_PATH\Source\Client\PA Shared\paAboutDialog.pas" -Find "PA_COPYRIGHT_YEAR.*= *'\d\d\d\d';" -Replace "PA_COPYRIGHT_YEAR = '$CURRENT_YEAR';"  # Set PA_COPYRIGHT_YEAR
                        }
                    }
                }
            } else {
                if ("$IS_2_TIER" -eq "TRUE") {  # 2 tier
                    if (Test-Path "$BUILD_TEMP_PATH\Source\Framework\AboutDialog.pas") {  # If source\Framework\AboutDialog.pas file exists
                        $VAR_RESULT = Find-InFile -Path "$BUILD_TEMP_PATH\Source\Framework\AboutDialog.pas" -Find "PA_COPYRIGHT_YEAR.*= *'$CURRENT_YEAR';"  # Set PA_COPYRIGHT_YEAR
                        if ("$VAR_RESULT" -le "1") {  # Only modify this file if the current year does not match
                            Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/framework/AboutDialog.pas" -Host "$VAULT_SERVER_ADDRESS"  # 
                            Replace-InFile -Path "$BUILD_TEMP_PATH\Source\Framework\AboutDialog.pas" -Find "PA_COPYRIGHT_YEAR.*= *'\d\d\d\d';" -Replace "PA_COPYRIGHT_YEAR = '$CURRENT_YEAR';"  # Set PA_COPYRIGHT_YEAR
                        }
                    }
                } else {  # 3 tier
                    if (Test-Path "$BUILD_TEMP_PATH\Source\Client\Framework\AboutDialog.pas") {  # If source\Client\Framework\AboutDialog.pas file exists
                        $VAR_RESULT = Find-InFile -Path "$BUILD_TEMP_PATH\Source\Client\Framework\AboutDialog.pas" -Find "PA_COPYRIGHT_YEAR.*= *'$CURRENT_YEAR';"  # Set PA_COPYRIGHT_YEAR
                        if ("$VAR_RESULT" -le "1") {  # Only modify this file if the current year does not match
                            Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/source/client/framework/AboutDialog.pas" -Host "$VAULT_SERVER_ADDRESS"  # 
                            Replace-InFile -Path "$BUILD_TEMP_PATH\Source\Client\Framework\AboutDialog.pas" -Find "PA_COPYRIGHT_YEAR.*= *'\d\d\d\d';" -Replace "PA_COPYRIGHT_YEAR = '$CURRENT_YEAR';"  # Set PA_COPYRIGHT_YEAR
                        }
                    }
                }
            }
            #endregion Update About dialog's copyright year if needed


            #region Check copyright year in EULA
            Write-Log "--- Check copyright year in EULA ---"
            $EULA_COPYRIGHT = Get-Content -Path "\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files\EULA.txt" -Raw  # Read \\%VAULT_SERVER_ADDRESS%\Groups\SDG\Setup Include Files\EULA.txt
            $EULA_COPYRIGHT = Get-SubstringBetween -Input "$EULA_COPYRIGHT" -Start "Professional Advantage Pty Ltd 1997 - " -End " All title and copyrights in and to the SOFTWARE PRODUCT"  # TODO: maybe just search for current year?
            if ("$EULA_COPYRIGHT" -ne "$CURRENT_YEAR") {
                Write-Log "[MESSAGE] "
            }
            Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files\EULA.txt" -Destination "$BUILD_TEMP_PATH\Setup" -Force  # Copy EULA to %BUILD_TEMP_PATH%\Setup
            #endregion Check copyright year in EULA

            Invoke-UpdateDccConfigFile   # Run UpdateDccConfigFile macro
        }
    }
    #endregion Update version numbers


    #region Run PASQL scripts/tests (if any)
    Write-Log "--- Run PASQL scripts/tests (if any) ---"
    if ("$SOURCE_CONTROL_LABEL" -ceq "") {  # If SOURCE_CONTROL_LABEL not set
        if ("$INI_SECTION" -notlike "Collect 5*") {  # Collect 5 quirk - scripts clash with Collect 6
            if ("$SETUP_PASQL_FILE_NAME" -eq "") {  # If SETUP_PASQL_FILE_NAME not yet set (nightly build)
                $VAR_RESULT = 0
                foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\scripts\*setup.pasql" -ErrorAction SilentlyContinue)) {  # Check for presence of *setup.pasql file
                    $SETUP_PASQL_FILE_NAME = $__file.FullName
                    $VAR_RESULT++
                }
            }
            if (("$VAR_RESULT" -ceq "1") -and ("$SETUP_PASQL_FILE_NAME" -ne "")) {  # If *setup.pasql file found
                if (Test-Path "$SETUP_PASQL_FILE_NAME.test") {  # If SETUP_PASQL_FILE_NAME.test exists
                    $SETUP_PASQL_FILE_NAME = "$SETUP_PASQL_FILE_NAME.test"  # Use it in tests
                }
                $TEMP_VAR = ""  # Clear TEMP_VAR
                if (Test-Path "$PA_UNIT_FILE_NAME") {  # Check if PA_UNIT_FILE_NAME exists

                    #region Check that PAUnitCMD is not older than MockDatabaseSettings.pas
                    Write-Log "--- Check that PAUnitCMD is not older than MockDatabaseSettings.pas ---"
                    if ("$NIGHTLY_BUILD" -ne "TRUE") {  # If not nightly build
                        $PA_UNIT_DATE = (Get-Item "$PA_UNIT_FILE_NAME").LastWriteTime.ToString()  # Get modification date of PAUnitCMD.exe
                        $MOCK_SETTINGS_DATE = Invoke-VaultCommand -Repository "SDG" -Command "LISTOBJECTPROPERTIES" -Parameters "`"$/PA Internal/PA Unit Testing/Trunk/Source/Framework/Test/MockDatabaseSettings.pas`""  # Get MockDatabaseSettings.pas modification date
                        $MOCK_SETTINGS_DATE = Get-SubstringBetween -Input "$MOCK_SETTINGS_DATE" -Start "<modifieddate>" -End "</modifieddate>"  # Extract date
                        # Script block (DelphiScript): Compare dates
                        $PAUNIT_TOO_OLD = "FALSE"
                        $paUnitDate = [datetime]::MinValue
                        $mockSettingsDate = [datetime]::MinValue
                        if ([datetime]::TryParse("$PA_UNIT_DATE", [ref]$paUnitDate) -and [datetime]::TryParse("$MOCK_SETTINGS_DATE", [ref]$mockSettingsDate)) {
                            if ($paUnitDate -lt $mockSettingsDate) {
                                $PAUNIT_TOO_OLD = "TRUE"
                            }
                        }
                        if ("$PAUNIT_TOO_OLD" -eq "TRUE") {  # If PAUnitCMD too old
                            $CONTINUE_WITH_OLD = Confirm-Action -Message "PAUnitCMD.exe is older than the MockDatabaseSettings.pas file and should be rebuilt first (PA Unit Testing project). PAUnitCMD.exe date: $PA_UNIT_DATE MockDatabaseSettings.pas date: $MOCK_SETTINGS_DATE  Continue with the old PAUnitCMD.exe?" -Default "True"  # 
                            if ("$CONTINUE_WITH_OLD" -ne "TRUE") {
                                throw "PAUnitCMD.exe is older than the MockDatabaseSettings.pas file and needs to be rebuilt. PAUnitCMD.exe date: $PA_UNIT_DATE MockDatabaseSettings.pas date: $MOCK_SETTINGS_DATE"
                            }
                        }
                    }
                    #endregion Check that PAUnitCMD is not older than MockDatabaseSettings.pas

                    Copy-FileEx -Source "$PA_UNIT_FILE_NAME" -Destination "$BUILD_TEMP_PATH\scripts" -Force  # Copy PAUnitCMD.exe
                    Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Setup Include Files\dbExpress\dbexpsda4*.dll" -Destination "$BUILD_TEMP_PATH\scripts" -Force  # Copy dbExpress dlls
                    $PA_UNIT_TESTS_FILE_NAME = "$BUILD_TEMP_PATH\scripts\PAUnitCMD.exe"  # Set PA_UNIT_TESTS_FILE_NAME
                } else {
                    throw "Unable to locate $PA_UNIT_FILE_NAME"
                }
                $REMOVE_UT_PASQ_FILE_NAME = "$BUILD_TEMP_PATH\scripts\remove_ut.remove"  # Set REMOVE_UT_PASQ_FILE_NAME
                Set-Content -Path "$REMOVE_UT_PASQ_FILE_NAME" -Value "script `"UT_REMOVE`", `"$Revision`: 1 $`", `"UT REMOVE`"    usedatabase(`"all`")    if exists(`"table`", `"PA_UT_REGISTERED_TESTS`") then     batch `"Truncate PA_UT_REGISTERED_TESTS table`"       sql `"         truncate table PA_UT_REGISTERED_TESTS       `"     end batch   end if  end script "  # Create REMOVE_UT_PASQ_FILE_NAME file
                $REMOVE_UT_PASQ_FILE_NAME = "<ROW DISPLAY_NAME=`"Remove UT`" SETUP_SCRIPT=`"$REMOVE_UT_PASQ_FILE_NAME`" RUN_ORDER=`"0`" SELECTED=`"Y`"/>"  # Set REMOVE_UT_PASQ_FILE_NAME to XML format
                $REMOVE_UT_PASQ_FILE_NAME = $REMOVE_UT_PASQ_FILE_NAME.Replace('\', '&#092;')  # Replace \ in REMOVE_UT_PASQ_FILE_NAME
                $SETUP_PASQL_FILE_NAME = "<ROW DISPLAY_NAME=`"Setup`" SETUP_SCRIPT=`"$SETUP_PASQL_FILE_NAME`" RUN_ORDER=`"1`" SELECTED=`"Y`"/>"  # Set SETUP_PASQL_FILE_NAME to XML format
                $SETUP_PASQL_FILE_NAME = $SETUP_PASQL_FILE_NAME.Replace('\', '&#092;')  # Replace \ in SETUP_PASQL_FILE_NAME
                $VAR_RESULT = 0
                foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\scripts\*setupTests.pasql" -ErrorAction SilentlyContinue)) {  # Check for presence of *setupTests.pasql file
                    $TEST_PASQL_FILE_NAME = $__file.FullName
                    $VAR_RESULT++
                }
                if ("$VAR_RESULT" -ceq "1") {  # If *setupTests.pasql file found
                    $TEST_PASQL_FILE_NAME = "<ROW DISPLAY_NAME=`"Tests`" SETUP_SCRIPT=`"$TEST_PASQL_FILE_NAME`" RUN_ORDER=`"2`" SELECTED=`"Y`"/>"  # Set TEST_PASQL_FILE_NAME to XML format
                    $TEST_PASQL_FILE_NAME = $TEST_PASQL_FILE_NAME.Replace('\', '&#092;')  # Replace \ in TEST_PASQL_FILE_NAME
                } else {
                    $TEST_PASQL_FILE_NAME = " "  # Clear TEST_PASQL_FILE_NAME
                }
                $SCRIPTS_XML_FILE_NAME = "$BUILD_TEMP_PATH\scripts\Scripts.xml"  # Set SCRIPTS_XML_FILE_NAME
                Remove-ItemSafe -Path "$BUILD_TEMP_PATH\scripts\*.xml"  # Delete old XML files
                if (Test-Path "$BUILD_TEMP_PATH\scripts\UnitTestingFramework.pasql") {  # Check if UnitTestingFramework.pasql exists
                    Set-ItemProperty -Path "$BUILD_TEMP_PATH\scripts\UnitTestingFramework.pasql" -Name IsReadOnly -Value $false  # Make writable
                    Replace-InFile -Path "$BUILD_TEMP_PATH\scripts\UnitTestingFramework.pasql" -Find "registertest\s*\(\s*`"PA_UT_MOCK_SQL_ERROR_UT`"\,\s*`"MOCK`"\,\s*`"MockSuite\.SubSuite`"\s*\)" -Replace "// registertest(`"PA_UT_MOCK_SQL_ERROR_UT`", `"MOCK`", `"MockSuite.SubSuite`")"  # Comment out tests that are designed to fail
                }
                $ORACLE = "/SUPPORTS_ORACLE:$SUPPORTS_ORACLE"  # Set ORACLE
                $SUN4 = "/SUPPORTS_SUN4:$SUPPORTS_SUN4"  # Set SUN4
                $SUN5 = "/SUPPORTS_SUN5:$SUPPORTS_SUN5"  # Set SUN5
                $SUN6 = "/SUPPORTS_SUN6:$SUPPORTS_SUN6"  # Set SUN6
                Remove-ItemSafe -Path "$BUILD_TEMP_PATH\scripts\*.log"  # Delete old log files
                Set-Content -Path "$BUILD_TEMP_PATH\scripts\test.vbs" -Value "  set FileSys = CreateObject(`"Scripting.FileSystemObject`")   set WshShell = CreateObject(`"WScript.Shell`")      set WshScriptExec = WshShell.Exec(`"`"`"$PA_UNIT_TESTS_FILE_NAME`"`" /SCRIPT_FILE:`"`"$SCRIPTS_XML_FILE_NAME`"`" /LOG_FILE:`"`"$BUILD_TEMP_PATH\scripts\nightly_build.log`"`" $ORACLE $SUN4 $SUN5 $SUN6`")    while WshScriptExec.Status = 0     while not WshScriptExec.StdOut.AtEndOfStream       Text = WshScriptExec.StdOut.Read(1)        WScript.StdOut.Write(Text)     wend   wend    WScript.Quit(WshScriptExec.ExitCode)"  # 
                Set-Content -Path "$SCRIPTS_XML_FILE_NAME" -Value "<?xml version=`"1.0`" standalone=`"yes`"?>   <DATAPACKET Version=`"2.0`">   <METADATA>     <FIELDS>       <FIELD attrname=`"DISPLAY_NAME`" fieldtype=`"string`" WIDTH=`"60`"/>       <FIELD attrname=`"SETUP_SCRIPT`" fieldtype=`"string`" WIDTH=`"255`"/>       <FIELD attrname=`"RUN_ORDER`" fieldtype=`"i4`"/>       <FIELD attrname=`"SELECTED`" fieldtype=`"string`" WIDTH=`"1`"/>     </FIELDS>     <PARAMS/>   </METADATA>   <ROWDATA>     $REMOVE_UT_PASQ_FILE_NAME   </ROWDATA> </DATAPACKET> "  # Create Scripts.xml file
                $VAR_RESULT = Invoke-Program -Path "cmd.exe" -Arguments "/c cscript.exe //NOLOGO `"$BUILD_TEMP_PATH\scripts\test.vbs`"" -WorkingDirectory "$BUILD_TEMP_PATH\scripts"  # Remove Unit tests
                Set-Content -Path "$SCRIPTS_XML_FILE_NAME" -Value "<?xml version=`"1.0`" standalone=`"yes`"?>   <DATAPACKET Version=`"2.0`">   <METADATA>     <FIELDS>       <FIELD attrname=`"DISPLAY_NAME`" fieldtype=`"string`" WIDTH=`"60`"/>       <FIELD attrname=`"SETUP_SCRIPT`" fieldtype=`"string`" WIDTH=`"255`"/>       <FIELD attrname=`"RUN_ORDER`" fieldtype=`"i4`"/>       <FIELD attrname=`"SELECTED`" fieldtype=`"string`" WIDTH=`"1`"/>     </FIELDS>     <PARAMS/>   </METADATA>   <ROWDATA>     $SETUP_PASQL_FILE_NAME     $TEST_PASQL_FILE_NAME   </ROWDATA> </DATAPACKET> "  # Create Scripts.xml file
                $VAR_RESULT = Invoke-Program -Path "cmd.exe" -Arguments "/c cscript.exe //NOLOGO `"$BUILD_TEMP_PATH\scripts\test.vbs`"" -WorkingDirectory "$BUILD_TEMP_PATH\scripts"  # Install scripts and run any Unit tests
                if (-not $?) {  # 
                    throw "Exceeded timeout running $PA_UNIT_TESTS_FILE_NAME"
                }
                # [DISABLED] Delete File(s)
                if ("$VAR_RESULT" -ne "0") {  # Check for errors
                    foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\scripts\nightly_build.log" -ErrorAction SilentlyContinue)) {  # 
                        $LOG_FILE = $__file.FullName
                        $LOG_FILE_TEXT = Get-Content -Path "$LOG_FILE" -Raw  # 
                        $LOG_FILE_TEXT = Get-SubstringBetween -Input "$LOG_FILE_TEXT" -Start "Run Results Begin:" -End "Run Results End:"  # 
                        if ("$LOG_FILE_TEXT" -like "*Run failed with the following errors*") {
                            if (("$LOG_FILE_TEXT" -like "*A later version of*") -and ("$LOG_FILE_TEXT" -like "*has been installed to this database*")) {  # If a later version has been installed into the database (cannot run scripts or tests)
                                $CAN_RUN_TESTS = "NO"  # set CAN_RUN_TESTS = NO
                            } else {
                                # [DISABLED] Throw
                                throw "Failed PASQL testing:  $LOG_FILE_TEXT"
                            }
                        }
                    }
                }
                if (("$PROJECT_TITLE" -like "PA Unit Testing*") -and ("$CAN_RUN_TESTS" -eq "YES")) {  # PA Unit Testing quirk - only PA Unit Testing can run MOCK FAIL TESTS
                    if (Test-Path "$BUILD_TEMP_PATH\scripts\UnitTestingFramework.pasql") {  # Check if UnitTestingFramework.pasql exists
                        Replace-InFile -Path "$BUILD_TEMP_PATH\scripts\UnitTestingFramework.pasql" -Find "//\s*registertest\s*\(\s*`"PA_UT_MOCK_SQL_ERROR_UT`"\,\s*`"MOCK`"\,\s*`"MockSuite\.SubSuite`"\s*\)" -Replace "registertest(`"PA_UT_MOCK_SQL_ERROR_UT`", `"MOCK`", `"MockSuite.SubSuite`")"  # Uncomment out tests that are designed to fail
                        $VAR_RESULT = Invoke-Program -Path "$PA_UNIT_TESTS_FILE_NAME" -Arguments "/SCRIPT_FILE:`"$SCRIPTS_XML_FILE_NAME`" /LOG_FILE:`"$BUILD_TEMP_PATH\scripts\nightly_build_2.log`" /INSTALL_ONLY:Y $ORACLE $SUN4 $SUN5 $SUN6"  # Install scripts again
                        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\scripts\nightly_build_2.log"  # 
                    }
                }
                # [DISABLED] Stop Macro Execution
            }
        } else {
            if ("$NIGHTLY_BUILD" -ne "TRUE") {  # If not nightly build
                Write-Log "[MESSAGE] "
            }
        }
    }
    Invoke-VaultCheckIn -Repository "SDG" -Path "$BUILD_TEMP_PATH\scripts/*" -Comment "$LABEL"  # Check in PASQL scripts
    # [DISABLED] Run Submacro (Get shared scripts from DevOps) - Get shared scripts from DevOps
    Invoke-get_files_from_devops   # Get shared scripts from GitHub
    #endregion Run PASQL scripts/tests (if any)


    #region Build projects
    Write-Log "--- Build projects ---"
    if ("$SOURCE_CONTROL_LABEL" -ne "") {  # If SOURCE_CONTROL_LABEL is set
        Write-Log "[MESSAGE] "
    }
    $TEMP_VAR = ("$BUILD_TEMP_PATH").Replace('\', '\\')  # Fix path
    New-Item -ItemType Directory -Path "$TEMP_VAR\debug" -Force | Out-Null  # Debug folder
    Remove-ItemSafe -Path "$BUILD_TEMP_PATH\Source\*.log" -Recurse  # Delete LOG files
    Remove-ItemSafe -Path "$BUILD_TEMP_PATH\Source\*.dcu" -Recurse  # Delete temp files
    foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\source\*.dpr" -ErrorAction SilentlyContinue)) {  # Loop through DPR project files, project name in VAR_RESULT_TEXT
        $VAR_RESULT_TEXT = $__file.FullName
        Invoke-BuildProject   # Build project
        if ("$VAR_RESULT_TEXT" -like "*\PAUnit.dpr") {  # PAUnit quirk - also create console app
            $TEMP_VAR = [System.IO.Path]::ChangeExtension("$VAR_RESULT_TEXT", 'CMD.dpr')  # Rename PAUnit.dpr to PAUnitCMD.dpr
            Copy-FileEx -Source "$BUILD_TEMP_PATH\source\PAUnit.dpr" -Destination "$BUILD_TEMP_PATH\source\PAUnitCMD.dpr" -Force  # Copy PAUnit.dpr to PAUnitCMD.dpr
            Copy-FileEx -Source "$BUILD_TEMP_PATH\source\PAUnit.dproj" -Destination "$BUILD_TEMP_PATH\source\PAUnitCMD.dproj" -Force  # Copy PAUnit.dproj to PAUnitCMD.dproj
            Replace-InFile -Path "$BUILD_TEMP_PATH\source\PAUnitCMD.dproj" -Find "PAUnit.exe" -Replace "PAUnitCMD.exe"  # Add CMD to project files
            Invoke-BuildProject   # Build project
        }
    }
    Copy-FileEx -Source "$BUILD_TEMP_PATH\*.exe" -Destination "$BUILD_TEMP_PATH\standard_files" -Force  # Make backup of standard executables
    #endregion Build projects


    #region Run tests
    Write-Log "--- Run tests ---"
    if (("$CAN_RUN_TESTS" -eq "YES") -and ("$SOURCE_CONTROL_LABEL" -ceq "")) {  # If CAN_RUN_TESTS = YES & SOURCE_CONTROL_LABEL not set
        $EXCEPTION_MESSAGE = ""  # Clear EXCEPTION_MESSAGE
        foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\*server.exe" -ErrorAction SilentlyContinue)) {  # Find files ending with server.exe, into TEMP_VAR
            $TEMP_VAR = $__file.FullName
            # [DISABLED] Execute DOS Command (Un-register server (needed for tests)) - Un-register server (needed for tests)
            $VAR_RESULT_TEXT = Invoke-DosCommand -Command "`"$TEMP_VAR`" -regserver"  # Register server (needed for tests)
        }
        foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\*svr.exe" -ErrorAction SilentlyContinue)) {  # Find files ending with svr.exe, into TEMP_VAR
            $TEMP_VAR = $__file.FullName
            # [DISABLED] Execute DOS Command (Un-register server (needed for tests)) - Un-register server (needed for tests)
            $VAR_RESULT_TEXT = Invoke-DosCommand -Command "`"$TEMP_VAR`" -regserver"  # Register server (needed for tests)
        }
        foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\SrchSetTask.dll" -ErrorAction SilentlyContinue)) {  # Collect 5 quirk - register SrchSetTask.dll for tests
            $TEMP_VAR = $__file.FullName
            $VAR_RESULT_TEXT = Invoke-DosCommand -Command "regsvr32 /s `"$TEMP_VAR`""  # Register server (needed for tests)
        }
        foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\*tests.exe" -ErrorAction SilentlyContinue)) {  # Get Test files and run them
            $TEMP_VAR = $__file.FullName
            if ("$TEMP_VAR" -notlike "*PAAUTHTests.exe") {  # Can't run PAAuthTests at this stage TODO:

                #region TODO: finish this (still need to check for test failures)
                Write-Log "--- TODO: finish this (still need to check for test failures) ---"
                Set-Content -Path "$BUILD_TEMP_PATH\test.vbs" -Value "  set FileSys = CreateObject(`"Scripting.FileSystemObject`")   set WshShell = CreateObject(`"WScript.Shell`")      set objTextFile = FileSys.OpenTextFile(`"$TEMP_VAR`" & `".log`", 2, True)   set WshScriptExec = WshShell.Exec(`"$TEMP_VAR`")    while WshScriptExec.Status = 0     while not WshScriptExec.StdOut.AtEndOfStream       Text = WshScriptExec.StdOut.Read(1)        WScript.StdOut.Write(Text)       objTextFile.Write(Text)     wend   wend    WScript.Quit(WshScriptExec.ExitCode) "  # 
                try {  # 
                    $VAR_RESULT = Invoke-Program -Path "cmd.exe" -Arguments "/c cscript.exe //NOLOGO `"$BUILD_TEMP_PATH\test.vbs`"" -WorkingDirectory "$BUILD_TEMP_PATH"
                } catch {  # 
                    if ("$VAR_RESULT" -ceq "-2") {  # TODO: does timeout always return -2?
                        $EXCEPTION_MESSAGE = "Test did not complete in 1 hour"
                    }
                    # Script block (DelphiScript): set CURRENT_PROCESS from TEMP_VAR
                    $CURRENT_PROCESS = [System.IO.Path]::GetFileName("$TEMP_VAR")
                    $CURRENT_PROCESS_NAME = [System.IO.Path]::GetFileNameWithoutExtension("$CURRENT_PROCESS")
                    if (Get-Process -Name "$CURRENT_PROCESS_NAME" -ErrorAction SilentlyContinue) {  # 
                        Stop-Process -Name "$CURRENT_PROCESS_NAME" -Force -ErrorAction SilentlyContinue  # Terminate running test
                    }
                }
                $VAR_RESULT_TEXT = Get-Content -Path "$TEMP_VAR.log" -Raw  # 
                Remove-ItemSafe -Path "$BUILD_TEMP_PATH\test.vbs"  # 
                #endregion TODO: finish this (still need to check for test failures)

                if ("$VAR_RESULT" -ne "0") {  # Check for errors
                    throw "Failed to run test:  $VAR_RESULT_TEXT  $EXCEPTION_MESSAGE"
                }
                if (("$VAR_RESULT_TEXT" -like "*FAILURES!!!*") -and ("$VAR_RESULT_TEXT" -like "*An error has occurred during program execution*")) {  # Check for errors
                    throw "Failed to run test:  $VAR_RESULT_TEXT"
                }
                Remove-ItemSafe -Path "$TEMP_VAR"  # Delete test files so there is no need to exclude them later
            }
        }
    }
    #endregion Run tests

    #endregion Build Delphi projects


    #region Build setups
    Write-Log "--- Build setups ---"
    if ("$SOURCE_CONTROL_LABEL" -ne "") {  # If SOURCE_CONTROL_LABEL set
        $BUILD_SETUPS = Confirm-Action -Message "Build standard setup files?  * THIS IS NOT RECOMMENDED! * WISE script may pick up files that are more recent than the original build." -Default "False"  # 
        if ("$DELPHI_VERSION" -ceq "6") {  # Delphi 6
            foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\source\*.dof" -ErrorAction SilentlyContinue)) {  # Loop through DOF project files, project name in VAR_RESULT_TEXT
                $DPR_FILE = $__file.FullName
                if ("$DPR_FILE" -ne "") {
                    break
                }
            }
        } else {
            foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\source\*.dproj" -ErrorAction SilentlyContinue)) {  # Loop through DPROJ project files, project name in VAR_RESULT_TEXT
                $DPR_FILE = $__file.FullName
                if ("$DPR_FILE" -ne "") {
                    break
                }
            }
        }
        if ("$DPR_FILE" -ne "") {
            $DPR_FILE_DATE = (Get-Item "$DPR_FILE").LastWriteTime.ToString()  # 
        } else {
            $DPR_FILE_DATE = "$SOURCE_CONTROL_LABEL"
        }
        # Script block (DelphiScript): normalize DPR_FILE_DATE for touch.exe
        $parsedDprDate = [datetime]::MinValue
        if ([datetime]::TryParse("$DPR_FILE_DATE", [System.Globalization.CultureInfo]::CurrentCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedDprDate)) {
            $DPR_FILE_DATE = $parsedDprDate.ToString("yyyy-MM-ddTHH:mm:ss")
        } else {
            $DPR_FILE_DATE = ""
        }
        if ("$DPR_FILE_DATE" -ne "") {
            Invoke-DosCommand -Command "`"C:\Freeware\touch.exe`" -a -m -x -d$DPR_FILE_DATE `"$BUILD_TEMP_PATH\*.exe`"" -WorkingDirectory "$BUILD_TEMP_PATH"  # 
            Invoke-DosCommand -Command "`"C:\Freeware\touch.exe`" -a -m -x -d$DPR_FILE_DATE `"$BUILD_TEMP_PATH\debug\*.exe`"" -WorkingDirectory "$BUILD_TEMP_PATH"  # 
        }
        Write-Log "[MESSAGE] "
    }
    if ("$BUILD_SETUPS" -eq "TRUE") {
        Copy-FileEx -Source "$BUILD_TEMP_PATH\*.exe" -Destination "$BUILD_TEMP_PATH\setup" -Force  # Copy all required exe files into the setup folder
        Copy-FileEx -Source "$BUILD_TEMP_PATH\*.chm" -Destination "$BUILD_TEMP_PATH\setup" -Force  # Copy all required help files into the setup folder
        Copy-FileEx -Source "$BUILD_TEMP_PATH\*.dll" -Destination "$BUILD_TEMP_PATH\setup" -Force  # Copy all required dll files into the setup folder
        $TEMP_VAR = "0.0.0.0"  # set TEMP_VAR to 0.0.0.0
        $VAR_RESULT = 0
        foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\scripts\*setup.pasql" -ErrorAction SilentlyContinue)) {  # Check for presence of *setup.pasql file
            $item = $__file.FullName
            $VAR_RESULT++
        }
        if ("$VAR_RESULT" -ceq "1") {  # If *setup.pasql found
            $TEMP_VAR = Invoke-DosCommand -Command "`"C:\Compilers\RAD Studio 2007\CodeGear\Bin\grep`" -i `"AppDBVersion *= .*\d+.*`" *setup.pasql" -WorkingDirectory "$BUILD_TEMP_PATH\scripts"  # Get DB version number from pasql script
            $TEMP_VAR = Get-SubstringBetween -Input "$TEMP_VAR" -Start "= `"" -End "`""  # 
        }
        foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\setup\*.exe" -ErrorAction SilentlyContinue)) {  # Get all EXEs for the build
            $TEMP_VAR_3 = $__file.FullName
            # [DISABLED] Get File Version Info (Get file version of the EXE) - Get file version of the EXE
            Invoke-Program -Path "C:\work\BuildStudio\sigcheck64.exe" -Arguments "/accepteula -nobanner -n `"$TEMP_VAR_3`"" -WorkingDirectory "C:\work\BuildStudio"  # Get file version of the EXE
            # [DISABLED] Execute Program (Get file version of the EXE) - Get file version of the EXE
            # [DISABLED] String Reverse
            # [DISABLED] String Concatenation
            # [DISABLED] String Substring
            $TEMP_VAR_2 = ("$TEMP_VAR_2").Trim()  # 
            # [DISABLED] String Reverse
            # [DISABLED] String Quoting
            # Script block (VBScript): Set TEMP_VAR to the higher of TEMP_VAR and TEMP_VAR_2
            # Converted from DelphiScript: Set TEMP_VAR to the higher of TEMP_VAR and TEMP_VAR_2
            # Script block (VBScript): Version comparison
            if ((Compare-Versions $TEMP_VAR_2 $TEMP_VAR) -gt 0) {
              $TEMP_VAR = $TEMP_VAR_2
              Write-Log "TEMP_VAR updated to: $TEMP_VAR"
            }
        }
        if ("$TEMP_VAR" -ceq "") {
            throw "Failed to get version number: "
        } else {
        }
        $VAR_RESULT = 0
        foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\Setup\Standard Setup.wse" -ErrorAction SilentlyContinue)) {  # Check for presence of Standard Setup.wse file
            $item = $__file.FullName
            $VAR_RESULT++
        }
        if (("$VAR_RESULT" -ceq "1") -and ("$PROJECT_TITLE" -notlike "Bank Reconciliation*")) {  # If Standard Setup.wse found and not building Bank Reconciliation

            #region Prepare external files required for the build
            Write-Log "--- Prepare external files required for the build ---"
            $INSTALL_WISE_FILE_NAME = ""  # Initialise INSTALL_WISE_FILE_NAME to blank
            foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\Setup\Include - Install Files.wse" -ErrorAction SilentlyContinue)) {  # Check for presence of Include - Install Files.wse file
                $INSTALL_WISE_FILE_NAME = $__file.FullName
            }
            # [DISABLED] If ... Then (If found) - If found
            foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\Setup\*.wse" -ErrorAction SilentlyContinue)) {  # Replace invalid path in wse files
                $INSTALL_WISE_FILE_NAME = $__file.FullName
                Replace-InFile -Path "$INSTALL_WISE_FILE_NAME" -Find "C:\Program Files\Wise" -Replace "C:\Program Files (x86)\Wise"  # Replace C:\PROGRA~1\WISEIN~1 with C:\PROGRA~2\WISEIN~1
            }
            #endregion Prepare external files required for the build

            $TEMP_VAR_3 = "$PROJECT_TITLE Setup v$TEMP_VAR.exe"  # Set TEMP_VAR_3
            Write-Log "Build the final setup name into TEMP_VAR_3"
            # Script block (DelphiScript): Build the Wise install string into TEMP_VAR_2
            $setupPath = Join-Path "$BUILD_TEMP_PATH" "Setup"
            $currentYearForSetup = if ($CURRENT_YEAR) { $CURRENT_YEAR } else { (Get-Date).Year }
            Write-Log "$setupPath\"
            $TEMP_VAR_2 = "`"C:\Program Files (x86)\Wise Installation System\Wise32.exe`" /d_PA_VERSION_=`"$TEMP_VAR`" /d_PA_COPYRIGHT_YEAR_=`"$currentYearForSetup`" /c `"$setupPath\Standard Setup.wse`""
            New-Item -ItemType Directory -Path "$BUILD_TEMP_PATH\builds" -Force | Out-Null  # Create directory for the final builds
            # LABEL: EXE build  (GoTo not supported in PS - restructure logic)
            Invoke-DosCommand -Command "$TEMP_VAR_2" -WorkingDirectory "$BUILD_TEMP_PATH\setup"  # Run Wise script
            # [DISABLED] Click Window Button (Cancel "The compiler variable _WISE_ does not point..." dialog, if any) - Cancel "The compiler variable _WISE_ does not point..." dialog, if any
            # [DISABLED] Wait for File (Wait for wise completion) - Wait for wise completion
            if (Test-Path "$BUILD_TEMP_PATH\setup\Standard Setup.EXE") {  # Check if Wise setup completed
                Move-Item -Path "$BUILD_TEMP_PATH\setup\Standard Setup.EXE" -Destination "$BUILD_TEMP_PATH\builds\$TEMP_VAR_3" -Force  # 
                Invoke-SignFile   # Sign final setup file
            } else {
                throw "Failed build setup file:  $TEMP_VAR_3"
            }
            if (Test-Path "$BUILD_TEMP_PATH\debug") {  # Check if we have a debug folder
                Copy-FileEx -Source "$BUILD_TEMP_PATH\debug\*.exe" -Destination "$BUILD_TEMP_PATH\setup" -Force  # Copy all required exe files into the setup folder
                Rename-Item -Path "$BUILD_TEMP_PATH\debug" -NewName "debug_files" -Force  # Rename the debug folder to prevent looping
                $TEMP_VAR_3 = "$PROJECT_TITLE Setup v$TEMP_VAR - debug.exe"  # Set TEMP_VAR_3
                Write-Log "Build the final setup name into TEMP_VAR_3"
                # GOTO: exe_build  (GoTo not supported in PS - restructure as loop)
            }
            # [DISABLED] File Enumerator (Cleanup setup folder) - Cleanup setup folder
        }
    }
    #endregion Build setups

    #endregion Build Standard setups (release and debug)


    #region Build non Wise setup
    Write-Log "--- Build non Wise setup ---"
    $VAR_RESULT = 0
    foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\*x.inf" -ErrorAction SilentlyContinue)) {  # Check for existence of .inf file
        $INF_FILE_NAME = $__file.FullName
        $VAR_RESULT++
    }
    if ("$VAR_RESULT" -ne "") {  # Get ActiveX projects only if needed

        #region Group
        Write-Log "--- Group ---"

        #region Get Source Control source files
        Write-Log "--- Get Source Control source files ---"
        Invoke-VaultGetLatest -Repository "SDG" -Path "$/Non Delphi Projects/ActiveX/*" -LocalFolder "$BUILD_TEMP_PATH\ActiveX"  # Get PA Applications source
        Invoke-VaultGetLatest -Repository "SDG" -Path "$/Products/ActiveX Launcher/*" -LocalFolder "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher"  # GetActiveX Launcher source
        #endregion Get Source Control source files


        #region Build ASP.Net solution
        Write-Log "--- Build ASP.Net solution ---"
        Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files\EULA.txt" -Destination "$BUILD_TEMP_PATH\ActiveX\InstallAware\License.txt" -Force  # Get EULA
        if (Test-Path "$BUILD_TEMP_PATH\source\dcc64.cfg") {  # If dcc64.cfg exists
            $BUILD_TYPE = "Release|x64"  # Release|x64
        } else {
            $BUILD_TYPE = "Release|x86"  # Release|x86
        }
        Invoke-MSBuild -SolutionFile "$BUILD_TEMP_PATH\ActiveX\PAApplications\PAApplications.sln" -Configuration "$BUILD_TYPE" -CompilerVersion "12.0"  # Compile PA Applications solution
        # [DISABLED] Execute DOS Command (Compile WiX Installer (RELEASE)) - Compile WiX Installer (RELEASE)
        if ("$VAR_RESULT" -ne "0") {  # Check for errors
            throw "Failed to build solution:  $VAR_RESULT_TEXT"
        }
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\ActiveX\PAApplications\PAApplications\*.cs" -Recurse  # Delete ActiveX\PAApplications\PAApplications\*.cs files
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\ActiveX\PAApplications\PAApplications\bin\*.pdb"  # Delete ActiveX\PAApplications\PAApplications\bin\*.pdb files
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\ActiveX\PAApplications\PAApplications\obj" -Recurse  # Remove obj folder
        #endregion Build ASP.Net solution


        #region Build ActiveX Launcher
        Write-Log "--- Build ActiveX Launcher ---"
        $OCX_ROOT_NAME = ("$INF_FILE_NAME").Replace('x.inf', ' ')  # 
        $OCX_ROOT_NAME = $OCX_ROOT_NAME.Replace('$BUILD_TEMP_PATH\', ' ')  # 
        $OCX_ROOT_NAME = ("$OCX_ROOT_NAME").Trim()  # Set OCX_ROOT_NAME
        $OCX_GUID = Get-IniValue -Path "$INF_FILE_NAME" -Section "ActiveXForm" -Key "CLASS_ActiveFormX"  # Set OCX_GUID
        # Get File Version Info: Set CLIENT_EXE_VERSION - no keys specified

        #region Build OCX files
        Write-Log "--- Build OCX files ---"
        # TODO: Stuff below here needs reviewing once XE2 is available
        $LIBID = Get-IniValue -Path "$INF_FILE_NAME" -Section "ActiveXForm" -Key "LIBID"  # Set LIBID
        $IID_IActiveFormX = Get-IniValue -Path "$INF_FILE_NAME" -Section "ActiveXForm" -Key "IID_IActiveFormX"  # Set IID_IActiveFormX
        $DIID_IActiveFormXEvents = Get-IniValue -Path "$INF_FILE_NAME" -Section "ActiveXForm" -Key "DIID_IActiveFormXEvents"  # Set DIID_IActiveFormXEvents
        $CLASS_ActiveFormX = Get-IniValue -Path "$INF_FILE_NAME" -Section "ActiveXForm" -Key "CLASS_ActiveFormX"  # Set CLASS_ActiveFormX
        $TxActiveFormBorderStyle = Get-IniValue -Path "$INF_FILE_NAME" -Section "ActiveXForm" -Key "TxActiveFormBorderStyle"  # Set TxActiveFormBorderStyle
        $TxPrintScale = Get-IniValue -Path "$INF_FILE_NAME" -Section "ActiveXForm" -Key "TxPrintScale"  # Set TxPrintScale
        $TxMouseButton = Get-IniValue -Path "$INF_FILE_NAME" -Section "ActiveXForm" -Key "TxMouseButton"  # Set TxMouseButton
        $TxPopupMode = Get-IniValue -Path "$INF_FILE_NAME" -Section "ActiveXForm" -Key "TxPopupMode"  # Set TxPopupMode
        Set-IniValue -Path "$INF_FILE_NAME" -Section "${OCX_ROOT_NAME}x.ocx" -Key "clsid" -Value "{$CLASS_ActiveFormX}"  # 
        Move-Item -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\ActiveXLauncher.dpr" -Destination "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\${OCX_ROOT_NAME}X.dpr" -Force  # Move ActiveXLauncher.dpr
        Move-Item -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\ActiveXLauncher.dproj" -Destination "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\${OCX_ROOT_NAME}X.dproj" -Force  # Move ActiveXLauncher.dproj
        Move-Item -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\ActiveXLauncher.idl" -Destination "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\${OCX_ROOT_NAME}X.idl" -Force  # Move ActiveXLauncher.idl
        Move-Item -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\ActiveXLauncher_TLB.pas" -Destination "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\${OCX_ROOT_NAME}X_TLB.pas" -Force  # Move ActiveXLauncher_TLB.pas
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\ActiveXLauncher_TLB.pas"  # _TLB.pas will be generated, remove ActiveXLauncher_TLB.pas
        # [DISABLED] Move File(s) (Move ActiveXLauncher.res) - Move ActiveXLauncher.res
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\ActiveXLauncher.res"  # .res will be generated, remove ActiveXLauncher.res
        Replace-InFile -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\${OCX_ROOT_NAME}X.dpr" -Find "ActiveXLauncher" -Replace "${OCX_ROOT_NAME}X"  # Set DPR file values
        Replace-InFile -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\${OCX_ROOT_NAME}X.dproj" -Find "ActiveXLauncher" -Replace "${OCX_ROOT_NAME}X"  # Set DPROJ file values
        Replace-InFile -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\${OCX_ROOT_NAME}X.idl" -Find "F390265A-12B0-46BD-8187-46F838E88D75" -Replace "$TxPopupMode"  # Set IDL file values
        Replace-InFile -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\${OCX_ROOT_NAME}X_TLB.pas" -Find "0468FA7A-065A-4E7D-A821-382A37683641" -Replace "$CLASS_ActiveFormX"  # Set _TLB.pas file values
        Replace-InFile -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\Application\ActiveFormImpl.pas" -Find "ActiveXLauncher" -Replace "${OCX_ROOT_NAME}X"  # Set ActiveFormImpl.pas file values
        Invoke-DosCommand -Command "`"C:\Compilers\RAD Studio 2007\CodeGear\Bin\gentlb.exe`" -P ${OCX_ROOT_NAME}X.idl" -WorkingDirectory "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source"  # Generate TLB and _TLB.pas files
        Copy-FileEx -Source "$BUILD_TEMP_PATH\source\$ICON_FILE_NAME" -Destination "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\source" -Force  # Get project icon file
        Set-Content -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\${OCX_ROOT_NAME}X.rc" -Value "MAINICON ICON `"$ICON_FILE_NAME`"  1 VERSIONINFO  FILEVERSION $V_MAJ,$V_MIN,$V_REL,$V_REV  PRODUCTVERSION $V_MAJ,$V_MIN,$V_REL  FILEFLAGSMASK 0x3fL  FILEFLAGS 0x0L  FILEOS 0x4L  //VFT_DLL (0x2L)  FILETYPE 0x2L  //VFT_APP  //FILETYPE 0x1L  FILESUBTYPE 0x0L BEGIN     BLOCK `"StringFileInfo`"     BEGIN         BLOCK `"0c0904e4`"         BEGIN             VALUE `"CompanyName`", `"Professional Advantage Pty. Ltd.`"             VALUE `"FileDescription`", `"$PROJECT_TITLE X`"             VALUE `"FileVersion`", `"$CLIENT_EXE_VERSION`"             VALUE `"LegalCopyright`", `"©1991 - $CURRENT_YEAR`"             VALUE `"ProductName`", `"$PROJECT_TITLE`"             VALUE `"ProductVersion`", `"$V_MAJOR.$V_MINOR.$V_RELEASE`"             VALUE `"OleSelfRegister`", `"1`"         END     END     BLOCK `"VarFileInfo`"     BEGIN         VALUE `"Translation`", 0xc09, 1252     END END "  # Create RC file
        Invoke-DosCommand -Command "`"C:\Program Files (x86)\Windows Kits\8.1\bin\x64\rc.exe`" /v /r /l0c09 /c1252 `"$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source\${OCX_ROOT_NAME}X.rc`"" -WorkingDirectory "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\Source"  # Build the RES file
        if (Test-Path "$BUILD_TEMP_PATH\source\dcc32.cfg") {  # If the main project has dcc32.cfg
            Remove-ItemSafe -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\source\dcc64.cfg"  # Delete dcc64.cfg config file
        } else {
            Remove-ItemSafe -Path "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\source\dcc32.cfg"  # Delete dcc32.cfg config file
        }
        Invoke-BuildProject   # Build project
        Copy-FileEx -Source "$BUILD_TEMP_PATH\ActiveX\ActiveX Launcher\${OCX_ROOT_NAME}X.ocx" -Destination "$BUILD_TEMP_PATH" -Force  # Copy OCX to the main folder
        #endregion Build OCX files


        #region Build CAB file
        Write-Log "--- Build CAB file ---"
        Set-IniValue -Path "$INF_FILE_NAME" -Section "Strings" -Key "DisplayLabel" -Value "$PROJECT_TITLE"  # Update INF file with the correct DisplayLabel value
        Replace-InFile -Path "$INF_FILE_NAME" -Find "^\x0D\x0A" -Replace ""  # Cleanup INF file
        $CABARC_COMMAND = " `"C:\Compilers\RAD Studio 2007\CodeGear\Bin\cabarc.exe`" n `"$BUILD_TEMP_PATH\${OCX_ROOT_NAME}ActiveX.cab`" `"$BUILD_TEMP_PATH\${OCX_ROOT_NAME}x.inf`""  # Prepare cabarc command  into CABARC_COMMAND
        foreach ($INF_FILE_NAME_ENTRY in (Get-IniSectionValues -Path "$INF_FILE_NAME" -Section "Add.Code")) {  # Loop through required files (listed in INF file)
            if ("$INF_FILE_NAME_ENTRY" -eq "midas.dll") {  # Construct full path and file name into FULL_FILE_NAME
                $FULL_FILE_NAME = "\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files\winsys\$INF_FILE_NAME_ENTRY"
            } else {
                $FULL_FILE_NAME = "$BUILD_TEMP_PATH\$INF_FILE_NAME_ENTRY"
            }
            $CABARC_COMMAND = "$CABARC_COMMAND" + ""  # Append FULL_FILE_NAME to CABARC_COMMAND
            # Get File Version Info: Get version of the current file into VER_NUM - no keys specified
            if ("$VER_NUM" -eq "") {  # If no version number exists use the EXE version
                $VER_NUM = "$CLIENT_EXE_VERSION"
            }
            $VER_NUM = $VER_NUM.Replace('.', ',')  # Replace dots with commas in version number
            Set-IniValue -Path "$BUILD_TEMP_PATH\${OCX_ROOT_NAME}x.inf" -Section "$INF_FILE_NAME_ENTRY" -Key "FileVersion" -Value "$VER_NUM"  # Update INF file with the VER_NUM
            # Path Manipulation: Get the file's extension
            # $FILE_TYPE = ... # TODO: verify path operation
            if ("$FILE_TYPE" -eq ".exe") {  # If the file is exe sign it
                Invoke-SignFile   # Sign file
            }
        }
        Write-Log "CABARC_COMMAND now holds cabarc command including parameters"
        $VAR_RESULT_TEXT = Invoke-DosCommand -Command "$CABARC_COMMAND"  # Create CAB file
        if ("$VAR_RESULT" -ne "0") {  # Check for errors
            throw "Failed to create CAB file:  $VAR_RESULT_TEXT"
        }
        Invoke-SignFile   # Sign CAB file
        #endregion Build CAB file

        #endregion Build ActiveX Launcher

        if ("$PROJECT_TITLE" -like "Bank Reconciliation*") {  # If building Bank Reconciliation

            #region Build Bank Reconciliation setup
            Write-Log "--- Build Bank Reconciliation setup ---"
            $PA_FRAMEWORK_FOLDER = "C:\Temp\Framework"  # Set PA_FRAMEWORK_FOLDER to C:\Temp\Framework (due to path too long error otherwise)
            try {  # 
                try {  # 
                    if (Test-Path "$PA_FRAMEWORK_FOLDER") {  # If directory exists
                        Invoke-DosCommand -Command "takeown /F $PA_FRAMEWORK_FOLDER /R"  # Take ownership of it
                    }
                    if (Test-Path "$PA_FRAMEWORK_FOLDER") {  # 
                        Remove-ItemSafe -Path "$PA_FRAMEWORK_FOLDER" -Recurse  # %PA_FRAMEWORK_FOLDER%
                    }
                } catch {  # 
                }
                if ("$SOURCE_CONTROL_LABEL" -ceq "") {  # If SOURCE_CONTROL_LABEL not set
                    Invoke-VaultGetLatest -Repository "SDG" -Path "$PA_VS_FRAMEWORK_PATH/*" -LocalFolder "$PA_FRAMEWORK_FOLDER"  # Get Framework Code into %PA_FRAMEWORK_FOLDER%
                } else {
                    Invoke-VaultGetByLabel -Repository "SDG" -Path "$PA_VS_FRAMEWORK_PATH/*" -Label "$SOURCE_CONTROL_LABEL" -LocalPath "$PA_FRAMEWORK_FOLDER"  # Get Framework Code by label  into %PA_FRAMEWORK_FOLDER%
                }
                # [DISABLED] Execute DOS Command (Download Framework packages) - Download Framework packages
                # [DISABLED] Execute DOS Command (Clean Framework (using msbuild 14.0 for now)) - Clean Framework (using msbuild 14.0 for now)
                # [DISABLED] Execute DOS Command (Compile Framework  (using msbuild 14.0 for now)) - Compile Framework  (using msbuild 14.0 for now)
                # [DISABLED] Execute DOS Command (Clean Framework) - Clean Framework
                # [DISABLED] Execute DOS Command (Compile Framework) - Compile Framework
                # [DISABLED] If ... Then (Check for errors) - Check for errors
                Replace-InFile -Path "$PA_FRAMEWORK_FOLDER\Compiled Resources\Server\CopyServerClientDll.bat" -Find "^pause" -Replace "echo done!"  # 
                Invoke-DosCommand -Command "$PA_FRAMEWORK_FOLDER\Compiled Resources\Server\CopyServerClientDll.bat`"" -WorkingDirectory "$PA_FRAMEWORK_FOLDER\Compiled Resources\Server"  # Copy resources
                # Script block (JScript): Set SOURCE_CONTROL_TRUNK_PATH
                $normalizedSourceControlPath = ("$SOURCE_CONTROL_SOURCE_PATH").Replace('\', '/')
                $SOURCE_CONTROL_TRUNK_PATH = [regex]::Replace("$normalizedSourceControlPath", "/delphi$", "", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                $PROJECT_LABEL_PATH = "$SOURCE_CONTROL_TRUNK_PATH"  # set PROJECT_LABEL_PATH
                # Script block (JScript): Set SOURCE_CONTROL_VISUAL_STUDIO_PATH
                $SOURCE_CONTROL_VISUAL_STUDIO_PATH = "$SOURCE_CONTROL_TRUNK_PATH/Visual Studio"
                # [DISABLED] Vault Cloak (Cloak SOURCE_CONTROL_SOURCE_PATH) - Cloak SOURCE_CONTROL_SOURCE_PATH
                $PA_FRAMEWORK_FOLDER = "$BUILD_TEMP_PATH\..\..\..\..\Framework\VisualStudio\PA.FrameWork\Trunk"  # Set PA_FRAMEWORK_FOLDER to %BUILD_TEMP_PATH%\..\..\..\..\Framework\VisualStudio\PA.FrameWork\Trunk
                if (Test-Path "$PA_FRAMEWORK_FOLDER") {  # %PA_FRAMEWORK_FOLDER%
                    try {  # 
                        Invoke-DosCommand -Command "takeown /F $PA_FRAMEWORK_FOLDER /R"  # Take ownership of it
                        Remove-ItemSafe -Path "$PA_FRAMEWORK_FOLDER" -Recurse  # 
                        Invoke-DosCommand -Command "mkdir $PA_FRAMEWORK_FOLDER"  # mkdir %PA_FRAMEWORK_FOLDER%
                    } catch {  # 
                    }
                }
                Copy-Item -Path "" -Destination "" -Recurse -Force  # Move C:\Temp\Framework to %PA_FRAMEWORK_FOLDER%
                # [DISABLED] Copy/Move Directory (Move PA.Blazor.Components to the correct relative location) - Move PA.Blazor.Components to the correct relative location
                # [DISABLED] If ... Then (If SOURCE_CONTROL_LABEL not set) - If SOURCE_CONTROL_LABEL not set
                # [DISABLED] Else
                # [DISABLED] Execute DOS Command (Download Framework Blazor packages) - Download Framework Blazor packages
                try {  # 
                    if (Test-Path "$BUILD_TEMP_PATH\..\Visual Studio") {  # If directory %BUILD_TEMP_PATH%\..\Visual Studio exists
                        Invoke-DosCommand -Command "takeown /F `"$BUILD_TEMP_PATH\..\Visual Studio`" /R"  # Take ownership of it
                    }
                    try {  # 
                        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\..\Visual Studio" -Recurse  # Delete old files if they exist
                    } catch {  # 
                        if (Get-Process -Name "Wise32.exe" -ErrorAction SilentlyContinue) {  # 
                            Stop-Process -Name "Wise32.exe" -Force -ErrorAction SilentlyContinue  # 
                            Remove-ItemSafe -Path "$BUILD_TEMP_PATH\..\Visual Studio" -Recurse  # Delete old files if they exist
                        }
                    }
                    Invoke-VaultGetLatest -Repository "SDG" -Path "$SOURCE_CONTROL_VISUAL_STUDIO_PATH/*" -LocalFolder "$BUILD_TEMP_PATH\..\Visual Studio"  # Get latest project source Visual Studio
                    Invoke-VaultGetLatest -Repository "SDG" -Path "$SOURCE_CONTROL_TRUNK_PATH/BankRecDllLoader/*" -LocalFolder "$BUILD_TEMP_PATH\..\BankRecDllLoader"  # Get latest project source BankRecDllLoader
                    Invoke-DosCommand -Command "`"$PA_FRAMEWORK_FOLDER\Source\.nuget\NuGet.exe`" restore `"$BUILD_TEMP_PATH\..\Visual Studio\BankReconciliation.sln`""  # Download project packages
                    # [DISABLED] Execute DOS Command (Download project packages) - Download project packages

                    #region Prepare various folders and paths for the build
                    Write-Log "--- Prepare various folders and paths for the build ---"
                    if (Test-Path "$BUILD_TEMP_PATH\Trunk") {  # 
                        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\Trunk" -Recurse  # %BUILD_TEMP_PATH%\Trunk
                    }
                    New-Item -ItemType Directory -Path "$BUILD_TEMP_PATH\Trunk\Delphi" -Force | Out-Null  # Create empty %BUILD_TEMP_PATH%\Trunk\Delphi
                    Invoke-DosCommand -Command "mklink /J  `"$BUILD_TEMP_PATH\Trunk\Delphi\ActiveX`" `"$BUILD_TEMP_PATH\ActiveX`""  # mklink /J  "%BUILD_TEMP_PATH%\Trunk\Delphi\ActiveX" "%BUILD_TEMP_PATH%\ActiveX"
                    # [DISABLED] Execute DOS Command (mklink /J  "%BUILD_TEMP_PATH%\Trunk\Delphi\Scripts" "%BUILD_TEMP_PATH%\Scripts") - mklink /J  "%BUILD_TEMP_PATH%\Trunk\Delphi\Scripts" "%BUILD_TEMP_PATH%\Scripts"
                    New-Item -ItemType Directory -Path "$BUILD_TEMP_PATH\Trunk\Delphi\Scripts" -Force | Out-Null  # %BUILD_TEMP_PATH%\Trunk\Delphi\Scripts
                    Copy-FileEx -Source "$BUILD_TEMP_PATH\Scripts\*.pasql" -Destination "$BUILD_TEMP_PATH\Trunk\Delphi\Scripts" -Force  # Copy %BUILD_TEMP_PATH%\Scripts\*.pasql to %BUILD_TEMP_PATH%\Trunk\Delphi\Scripts
                    Invoke-DosCommand -Command "mklink /J  `"$BUILD_TEMP_PATH\Trunk\Delphi\Setup`" `"$BUILD_TEMP_PATH\Setup`""  # mklink /J  "%BUILD_TEMP_PATH%\Trunk\Delphi\Setup" "%BUILD_TEMP_PATH%\Setup"
                    Invoke-DosCommand -Command "mklink /J  `"$BUILD_TEMP_PATH\Trunk\Delphi\Source`" `"$BUILD_TEMP_PATH\Source`""  # mklink /J  "%BUILD_TEMP_PATH%\Trunk\Delphi\Source" "%BUILD_TEMP_PATH%\Source"
                    Invoke-DosCommand -Command "mklink /J  `"$BUILD_TEMP_PATH\Trunk\Delphi\Reports`" `"$BUILD_TEMP_PATH\Reports`""  # mklink /J  "%BUILD_TEMP_PATH%\Trunk\Delphi\Reports" "%BUILD_TEMP_PATH%\Reports"
                    Copy-FileEx -Source "$BUILD_TEMP_PATH\*.exe" -Destination "$BUILD_TEMP_PATH\Trunk\Delphi" -Force  # Copy %BUILD_TEMP_PATH%\*.exe to %BUILD_TEMP_PATH%\Trunk\Delphi
                    Copy-FileEx -Source "$BUILD_TEMP_PATH\*.chm" -Destination "$BUILD_TEMP_PATH\Trunk\Delphi" -Force  # Copy %BUILD_TEMP_PATH%\*.chm to%BUILD_TEMP_PATH%\Trunk\Delphi
                    Copy-FileEx -Source "$BUILD_TEMP_PATH\*.cab" -Destination "$BUILD_TEMP_PATH\Trunk\Delphi" -Force  # Copy %BUILD_TEMP_PATH%\*.cab to %BUILD_TEMP_PATH%\Trunk\Delphi
                    Copy-FileEx -Source "$BUILD_TEMP_PATH\*.inf" -Destination "$BUILD_TEMP_PATH\Trunk\Delphi" -Force  # Copy %BUILD_TEMP_PATH%\*.inf to %BUILD_TEMP_PATH%\Trunk\Delphi
                    # [DISABLED] Throw (For testing Vault undo checkout) - For testing Vault undo checkout
                    #endregion Prepare various folders and paths for the build

                    if ("$NIGHTLY_BUILD" -ne "TRUE") {  # If not nightly build
                        # [DISABLED] Vault Check Out (AssemblyInfo.cs) - AssemblyInfo.cs
                        # [DISABLED] Edit Assembly Info (Configure AssemblyInfo.cs) - Configure AssemblyInfo.cs
                        Invoke-VaultCommand -Repository "SDG" -Command "SETWORKINGFOLDER" -Parameters "-forcesubfolderstoinherit  `"$SOURCE_CONTROL_VISUAL_STUDIO_PATH`" `"$BUILD_TEMP_PATH\..\Visual Studio`""  # Set working folder
                        Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_VISUAL_STUDIO_PATH/BankReconciliation/BankReconciliation.csproj" -Host "$VAULT_SERVER_ADDRESS"  # BankReconciliation.csproj
                        Replace-InFile -Path "$BUILD_TEMP_PATH\..\Visual Studio\BankReconciliation\BankReconciliation.csproj" -Find "<Version>\d+\.\d+\.\d+\.\d+<" -Replace "<Version>$BUILD_VERSION<"  # Set version number in BankReconciliation.csproj
                        Set-XmlValue -Path "" -XPath "//processing-instruction('define')[contains(.,'MajorVersion')]" -Value "MajorVersion=`"$V_MAJOR`" "  # Set MajorVersion in Parameters.wxi
                        Set-XmlValue -Path "" -XPath "//processing-instruction('define')[contains(.,'MinorVersion')]" -Value "MinorVersion=`"$V_MINOR`""  # Set MinorVersion in Parameters.wxi
                        Set-XmlValue -Path "" -XPath "//processing-instruction('define')[contains(.,'ReleaseVersion')]" -Value "ReleaseVersion=`"$V_RELEASE`""  # Set ReleaseVersion in Parameters.wxi
                        Set-XmlValue -Path "" -XPath "//processing-instruction('define')[contains(.,'BuildVersion')]" -Value "BuildVersion=`"$V_BUILD`""  # Set BuildVersion in Parameters.wxi
                    }
                    Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files\PA Utils\RTF Tool\TextToRtfConverter.exe" -Destination "$BUILD_TEMP_PATH\.." -Force  # Get TextToRtfConverter.exe
                    $VAR_RESULT_TEXT = Invoke-DosCommand -Command "`"$BUILD_TEMP_PATH\..\TextToRtfConverter.exe`" `"$BUILD_TEMP_PATH\Setup\EULA.txt`" `"$BUILD_TEMP_PATH\..\Visual Studio\BankReconciliationSetup\EULA.rtf`""  # Convert Eula to RTF format
                    # [DISABLED] Execute DOS Command (Download Bank Rec packages) - Download Bank Rec packages
                    Invoke-DosCommand -Command "npm i" -WorkingDirectory "$BUILD_TEMP_PATH\..\Visual Studio\BankReconciliation\ClientApp"  # Restore npm packages
                    Invoke-DosCommand -Command "npm install" -WorkingDirectory "$BUILD_TEMP_PATH\..\Visual Studio\BankReconciliation\ClientApp"  # Restore ClientApp node_modules
                    if ("$BUILD_RESULT" -like "*error*") {  # Check for errors
                        throw "Failed to restore node_modules:  $BUILD_RESULT"
                    }
                    # [DISABLED] Execute DOS Command (Compile Bank Rec project (RELEASE)) - Compile Bank Rec project (RELEASE)
                    if (Test-Path "$BUILD_TEMP_PATH\..\Visual Studio\BankReconciliationHarvest") {  # If exists %BUILD_TEMP_PATH%\..\Visual Studio\BankReconciliationHarvest
                        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\..\Visual Studio\BankReconciliationHarvest" -Recurse  # %BUILD_TEMP_PATH%\..\Visual Studio\BankReconciliationHarvest
                    }
                    # [DISABLED] Execute DOS Command (Compile Bank Rec project (RELEASE)) - Compile Bank Rec project (RELEASE)
                    # [DISABLED] Execute DOS Command (Compile Bank Rec project (RELEASE)) - Compile Bank Rec project (RELEASE)
                    # [DISABLED] Execute DOS Command (Compile Bank Rec project (RELEASE)) - Compile Bank Rec project (RELEASE)
                    # [DISABLED] If ... Then (Check for errors) - Check for errors
                    if (Test-Path "$BUILD_TEMP_PATH\..\Visual Studio\Resources") {  # 
                    } else {
                        New-Item -ItemType Directory -Path "$BUILD_TEMP_PATH\..\Visual Studio\Resources" -Force | Out-Null  # create \..\Visual Studio\Resources
                    }
                    Copy-FileEx -Source "$PA_FRAMEWORK_FOLDER\Compiled Resources\Dependencies\PAAuthClient32.dll" -Destination "$BUILD_TEMP_PATH\..\Visual Studio\Resources" -Force  # copy PAAuthClient32.dll
                    Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files\PA PGP Utils\*.dll" -Destination "$BUILD_TEMP_PATH\..\Visual Studio\Resources" -Force  # copy PGP Utils dlls
                    Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files\dbExpress\dbexpsda40.dll" -Destination "$BUILD_TEMP_PATH\..\Visual Studio\Resources" -Force  # copy dbexpsda40.dll
                    Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files\dbExpress\dbexpsda41.dll" -Destination "$BUILD_TEMP_PATH\..\Visual Studio\Resources" -Force  # copy dbexpsda41.dll
                    # [DISABLED] Execute DOS Command (Publish Bank Rec project (RELEASE, .NET core 2.1) TODO: use FolderProfile.pubxml config file once moved to 2.2) - Publish Bank Rec project (RELEASE, .NET core 2.1) TODO: use FolderProfile.pubxml config file once moved to 2.2
                    if (Test-Path "$BUILD_TEMP_PATH\..\Visual Studio\BankReconciliation\..\BankReconciliationHarvest") {  # 
                    } else {
                        New-Item -ItemType Directory -Path "$BUILD_TEMP_PATH\..\Visual Studio\BankReconciliation\..\BankReconciliationHarvest" -Force | Out-Null  # create \..\BankReconciliationHarvest
                    }
                    $BUILD_RESULT = Invoke-DosCommand -Command "$DOTNET publish -p:PublishProfile=FolderProfile --version-suffix $V_MAJOR.$V_MINOR.$V_RELEASE.$V_BUILD -o ..\BankReconciliationHarvest" -WorkingDirectory "$BUILD_TEMP_PATH\..\Visual Studio\BankReconciliation"  # Publish BankReconciliation project
                    # [DISABLED] If ... Then (Check for errors) - Check for errors
                    Invoke-DosCommand -Command "$MSBUILD_EXE `"$BUILD_TEMP_PATH\..\Visual Studio\BankReconciliationWix.sln`" `"/t:REBUILD`" `"/p:Configuration=Release`" `"/p:Platform=x86`""  # Compile WiX Installer (RELEASE)
                    # [DISABLED] If ... Then (Check for errors) - Check for errors
                    Move-Item -Path "$BUILD_TEMP_PATH\..\Visual Studio\BankReconciliationSetup\bin\Release\*Setup.msi" -Destination "$BUILD_TEMP_PATH\builds\$PROJECT_TITLE Setup v$BUILD_VERSION.msi" -Force  # 
                    Invoke-SignFile   # Sign MSI file
                    $VS_LABEL = "Build $BUILD_VERSION"  # Set VS_LABEL
                    Invoke-VaultCheckIn -Repository "SDG" -Path "$SOURCE_CONTROL_VISUAL_STUDIO_PATH/*" -Comment "$VS_LABEL"  # 
                } finally {  # 
                    # [DISABLED] Vault Uncloak (Uncloak SOURCE_CONTROL_SOURCE_PATH) - Uncloak SOURCE_CONTROL_SOURCE_PATH
                    Remove-ItemSafe -Path "$BUILD_TEMP_PATH\Trunk" -Recurse  # %BUILD_TEMP_PATH%\Trunk
                    # [DISABLED] Set/Reset Variable Value (Set SOURCE_CONTROL_SOURCE_PATH to SOURCE_CONTROL_TRUNK_SOURCE_PATH) - Set SOURCE_CONTROL_SOURCE_PATH to SOURCE_CONTROL_TRUNK_SOURCE_PATH
                    # [DISABLED] Vault Custom Command (Set working folder) - Set working folder
                }
                if (Test-Path "$PA_FRAMEWORK_FOLDER") {  # 
                    try {  # 
                        Remove-ItemSafe -Path "$PA_FRAMEWORK_FOLDER" -Recurse  # %PA_FRAMEWORK_FOLDER%
                    } catch {  # 
                    }
                }
            } catch {  # 
                Invoke-VaultUndoCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_VISUAL_STUDIO_PATH/*"  # Undo check out
                throw "$EXCEPTION_MESSAGE"
            }
            #endregion Build Bank Reconciliation setup

        } else {
            if (Test-Path "$BUILD_TEMP_PATH\Setup\Setup.sln") {  # If file Setup.sln exists (new setup) fall through to outer statement
            } else {  # (old setup)

                #region Build ActiveX setup exe
                Write-Log "--- Build ActiveX setup exe ---"
                $REVISION_GUID = [guid]::NewGuid().ToString('Plain').ToUpper()  # 
                Replace-InFile -Path "$BUILD_TEMP_PATH\ActiveX\InstallAware\PA Applications.mpr" -Find "`"TITLE=PA Applications`"" -Replace "`"TITLE=$PROJECT_TITLE ActiveX`""  # Set InstallAware project values
                Invoke-InstallAware -ProjectFile "$BUILD_TEMP_PATH\ActiveX\InstallAware\PA Applications.mpr" -BuildType "Compressed Single Self Installing EXE"  # Build Active X setup
                if ("$VAR_RESULT" -ne "0") {  # Check for errors
                    throw "Failed to build InstallAware project:  $VAR_RESULT_TEXT"
                }
                Invoke-SignFile   # Sign Active X setup file
                Copy-FileEx -Source "$BUILD_TEMP_PATH\ActiveX\InstallAware\Release\Single\*.exe" -Destination "$BUILD_TEMP_PATH\builds" -Force  # 
                #endregion Build ActiveX setup exe

            }
        }
        #endregion Group

    } else {
        $VAR_RESULT = 0
        foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\*wix.sln" -ErrorAction SilentlyContinue)) {  # Check for existence of *wix.sln file
            $INF_FILE_NAME = $__file.FullName
            $VAR_RESULT++
        }
        if ("$VAR_RESULT" -ne "") {  # Build solution if soltion file exists
            Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files\EULA.txt" -Destination "$BUILD_TEMP_PATH\Setup" -Force  # Get EULA
            Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files\PA Utils\RTF Tool\TextToRtfConverter.exe" -Destination "$BUILD_TEMP_PATH\.." -Force  # Get TextToRtfConverter.exe
            $VAR_RESULT_TEXT = Invoke-DosCommand -Command "`"$BUILD_TEMP_PATH\..\TextToRtfConverter.exe`" `"$BUILD_TEMP_PATH\Setup\EULA.txt`" `"$BUILD_TEMP_PATH\Setup\EULA.rtf`""  # Convert Eula to RTF format
            Invoke-DosCommand -Command "$MSBUILD_EXE `"$INF_FILE_NAME`" `"/t:REBUILD`" `"/p:Configuration=Release`" `"/p:Platform=x86`""  # Compile WiX Installer (RELEASE)
            New-Item -ItemType Directory -Path "$BUILD_TEMP_PATH\builds" -Force | Out-Null  # Create %BUILD_TEMP_PATH%\builds
            Move-Item -Path "$BUILD_TEMP_PATH\Setup\bin\Release\*Setup.msi" -Destination "$BUILD_TEMP_PATH\builds\$PROJECT_TITLE Setup v$BUILD_VERSION.msi" -Force  # 
            Invoke-SignFile   # Sign MSI file
        }
    }
    if (Test-Path "$BUILD_TEMP_PATH\Setup\Setup.sln") {  # If file Setup.sln exists (new setup)

        #region Build solution
        Write-Log "--- Build solution ---"
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\Setup\*.exe"  # Delete exes
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\Setup\*.dll"  # Delete dlls
        Invoke-VaultGetLatest -Repository "SDG" -Path "$/Framework/Icons/APAICON.ico" -LocalFolder "$BUILD_TEMP_PATH"  # Get APAICON.ico (needed by WIX for PAApplications)
        Copy-FileEx -Source "\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files\PA Utils\RTF Tool\TextToRtfConverter.exe" -Destination "$BUILD_TEMP_PATH\Setup" -Force  # Get TextToRtfConverter.exe
        $VAR_RESULT_TEXT = Invoke-DosCommand -Command "`"$BUILD_TEMP_PATH\Setup\TextToRtfConverter.exe`" `"$BUILD_TEMP_PATH\Setup\EULA.txt`" `"$BUILD_TEMP_PATH\Setup\EULA.rtf`""  # Convert Eula to RTF format
        Invoke-VaultCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/Setup/ApplicationConstants.cs" -Host "$VAULT_SERVER_ADDRESS"  # Check out version file
        $VERSION = "$BUILD_VERSION"
        Replace-InFile -Path "$BUILD_TEMP_PATH\Setup\ApplicationConstants.cs" -Find "VersionNumber\s*=\s*\`"[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+\`";" -Replace "VersionNumber = \`"$VERSION\`";"  # Set project version number
        Invoke-DosCommand -Command "$NUGET_EXE restore `"Setup.sln`"" -WorkingDirectory "$BUILD_TEMP_PATH\Setup"  # Download nuget  packages (older)
        Invoke-DosCommand -Command "dotnet restore `"Setup.sln`" -r win7-x86 /p:Configuration=Release" -WorkingDirectory "$BUILD_TEMP_PATH\Setup"  # Download nuget  packages (newer)
        # [DISABLED] If Directory Exists (%BUILD_TEMP_PATH%\SetupTests) - %BUILD_TEMP_PATH%\SetupTests
        New-Item -ItemType Directory -Path "$BUILD_TEMP_PATH\builds" -Force | Out-Null  # Create %BUILD_TEMP_PATH%\builds
        if (Test-Path "$BUILD_TEMP_PATH\debug_files") {  # If the debug folder was renamed by the Wise setup macro
            Rename-Item -Path "$BUILD_TEMP_PATH\debug_files" -NewName "debug" -Force  # Rename it back
        }
        # LABEL: wix build  (GoTo not supported in PS - restructure logic)
        Invoke-DosCommand -Command "$MSBUILD_2019_EXE `"Setup.sln`" `"/t:REBUILD`" `"/p:Configuration=Release`"" -WorkingDirectory "$BUILD_TEMP_PATH\Setup"  # Compile WiX Installer (RELEASE)
        $MSI_FILE = ""
        foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\Setup\*.msi" -ErrorAction SilentlyContinue)) {  # Check for MSI files
            $MSI_FILE = $__file.FullName
        }
        if ("$MSI_FILE" -ne "") {  # If MSI file found
            $MSI_FILE = Split-Path -Path "$MSI_FILE" -Leaf  # Extract file name only
            if (Test-Path "$BUILD_TEMP_PATH\builds\$MSI_FILE") {  # If the file already exists in the builds folder
                $NEW_MSI_FILE = ("$MSI_FILE").Replace('.msi', ' - debug.msi')  # Replace ".msi" with  "- debug.msi"
            } else {
                $NEW_MSI_FILE = "$MSI_FILE"  # Set new msi file name
            }
            Move-Item -Path "$BUILD_TEMP_PATH\Setup\$MSI_FILE" -Destination "$BUILD_TEMP_PATH\builds\$NEW_MSI_FILE" -Force  # This is the release setup
            Invoke-SignFile   # Sign MSI file
        } else {
            throw "`"$BUILD_TEMP_PATH\Setup\Setup.sln`" failed to build!"
        }
        if (Test-Path "$BUILD_TEMP_PATH\debug") {  # Check if we have a debug folder
            Copy-FileEx -Source "$BUILD_TEMP_PATH\debug\*.exe" -Destination "$BUILD_TEMP_PATH" -Force  # Copy all required exe files into the setup folder
            Copy-FileEx -Source "$BUILD_TEMP_PATH\debug\*.dll" -Destination "$BUILD_TEMP_PATH" -Force  # Copy all required dlls files into the setup folder
            Rename-Item -Path "$BUILD_TEMP_PATH\debug" -NewName "debug_files" -Force  # Rename the debug folder to prevent looping
            # GOTO: wix_build_1  (GoTo not supported in PS - restructure as loop)
        }
        #endregion Build solution

    }
    #endregion Build non Wise setup

    if ("$NIGHTLY_BUILD" -ne "TRUE") {  # If not nightly build

        #region Move all setups to the release folder and commit Source Control changes
        Write-Log "--- Move all setups to the release folder and commit Source Control changes ---"

        #region Move all setups to the release folder
        Write-Log "--- Move all setups to the release folder ---"
        $TEMP_VAR = "\\$VAULT_SERVER_ADDRESS\ForQA\$PROJECT_TITLE\$RELEASE_VERSION"  # Final destination
        if ("$PROJECT_TITLE" -like "Contract and Service Billing*") {  # CSB quirk - release folder does not match project title (PROJECT_TITLE)
            $TEMP_VAR = "\\$VAULT_SERVER_ADDRESS\ForQA\CSB\$RELEASE_VERSION"  # Final destination
        }
        if ("$PROJECT_TITLE" -like "Pillar SunSystems Interface*") {  # Pillar SunSystems Interface quirk - release folder does not match project title (PROJECT_TITLE)
            $TEMP_VAR = "\\$VAULT_SERVER_ADDRESS\ForQA\Pillar\SunSystems Interface\$RELEASE_VERSION"  # Final destination
        }
        if ("$PROJECT_TITLE" -like "Pillar Member and Sundry Payments*") {  # Pillar Member and Sundry Payments quirk - release folder does not match project title (PROJECT_TITLE)
            $TEMP_VAR = "\\$VAULT_SERVER_ADDRESS\ForQA\Pillar\Member and Sundry Payments\$RELEASE_VERSION"  # Final destination
        }
        if ("$PROJECT_TITLE" -like "Phone Billing*") {  # Phone Billing quirk - release folder does not match project title (PROJECT_TITLE)
            $TEMP_VAR = "\\$VAULT_SERVER_ADDRESS\ForQA\PA Internal\Phone Billing"  # Final destination
        }
        if ("$PROJECT_TITLE" -like "FuturePlus Receipt Registration*") {  # FuturePlus Receipt Registration quirk - release folder does not match project title (PROJECT_TITLE)
            $TEMP_VAR = "\\$VAULT_SERVER_ADDRESS\forqa\FuturePlus\$PROJECT_TITLE\$RELEASE_VERSION"  # Final destination
        }
        if ("$PROJECT_TITLE" -eq "FuturePlus") {  # FuturePlus Classic Interface quirk - release folder does not match project title (PROJECT_TITLE)
            $TEMP_VAR = "\\$VAULT_SERVER_ADDRESS\forqa\$PROJECT_TITLE\Classic Interface\$RELEASE_VERSION"  # Final destination
        }
        if ("$INI_SECTION" -like "ePay Japanese*") {  # ePay Japanese quirk - release folder does not match project title (PROJECT_TITLE)
            $TEMP_VAR = "\\$VAULT_SERVER_ADDRESS\ForQA\$PROJECT_TITLE\$RELEASE_VERSION\Japanese"  # Final destination
        }
        if ("$INI_SECTION" -like "PA Documents Viewer*") {  # PA Documents Viewer quirk - release folder is \\%VAULT_SERVER_ADDRESS%\Groups\SDG\Setup Include Files\PA Documents Viewer
            $TEMP_VAR = "\\$VAULT_SERVER_ADDRESS\Groups\SDG\Setup Include Files\PA Documents Viewer"  # Final destination
        }
        if ("$SOURCE_CONTROL_LABEL" -ne "") {  # If SOURCE_CONTROL_LABEL is set
            $TEMP_VAR = "$TEMP_VAR\Rebuild"  # Final destination
        }
        if (Test-Path "$TEMP_VAR") {  # Check if \\%VAULT_SERVER_ADDRESS%\FORQA... directory exists
        } else {
            New-Item -ItemType Directory -Path "$TEMP_VAR" -Force | Out-Null  # 
        }
        $OVERWRITE_FILES = "false"  # Set OVERWRITE_FILES to false
        if (Test-Path "$BUILD_TEMP_PATH\builds") {  # Check if builds... directory exists
            foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\builds\*.*" -ErrorAction SilentlyContinue)) {  # 
                $TEMP_VAR_2 = $__file.FullName
                $TEMP_VAR_3 = Split-Path -Path "$TEMP_VAR_2" -Leaf  # 
                if (Test-Path "$TEMP_VAR\$TEMP_VAR_3") {  # 
                    $OVERWRITE_FILES = Confirm-Action -Message "File $TEMP_VAR\$TEMP_VAR_3 already exists. Overwrite existing files?" -Default "True"  # 
                    break
                }
            }
            if (("$OVERWRITE_FILES" -eq "true") -and ("$OVERWRITE_FILES" -eq "yes")) {  # If OVERWRITE_FILES = true
                Copy-FileEx -Source "$BUILD_TEMP_PATH\builds\*.*" -Destination "$TEMP_VAR" -Force  # Overwrite copy files to \\%VAULT_SERVER_ADDRESS%\forqa\...
                if ("$VAR_RESULT" -ceq "0") {  # Check for errors
                    Write-Log "[MESSAGE] "
                } else {
                }
            } else {
                Copy-FileEx -Source "$BUILD_TEMP_PATH\builds\*.*" -Destination "$TEMP_VAR"  # No overwrite copy files to \\%VAULT_SERVER_ADDRESS%\forqa\...
                if ("$VAR_RESULT" -ceq "0") {  # Check for errors
                    Write-Log "[MESSAGE] "
                } else {
                }
            }
        } else {
            if (Test-Path "$BUILD_TEMP_PATH\setup") {  # Check if Setup... directory exists
                foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\setup\*.exe" -ErrorAction SilentlyContinue)) {  # 
                    $TEMP_VAR_2 = $__file.FullName
                    $TEMP_VAR_3 = Split-Path -Path "$TEMP_VAR_2" -Leaf  # 
                    if (Test-Path "$TEMP_VAR\$TEMP_VAR_3") {  # 
                        if ("$INI_SECTION" -notlike "PA Documents Viewer*") {  # PA Documents Viewer quirk - only prompt if not PA Document Viewer
                            $OVERWRITE_FILES = Confirm-Action -Message "File $TEMP_VAR\$TEMP_VAR_3 already exists. Overwrite existing files?" -Default "False"  # 
                            break
                        } else {
                            $OVERWRITE_FILES = "true"  # Set OVERWRITE_FILES to true
                        }
                    }
                }
                if ("$OVERWRITE_FILES" -eq "true") {  # If OVERWRITE_FILES = true
                    if ("$INI_SECTION" -like "PA Documents Viewer*") {  # PA Documents Viewer quirk - only copy exes
                        Copy-FileEx -Source "$BUILD_TEMP_PATH\setup\*.exe" -Destination "$TEMP_VAR" -Force  # Overwrite copy files to \\%VAULT_SERVER_ADDRESS%\forqa\...
                    } else {
                        Copy-FileEx -Source "$BUILD_TEMP_PATH\setup\*.*" -Destination "$TEMP_VAR" -Force  # Overwrite copy files to \\%VAULT_SERVER_ADDRESS%\forqa\...
                    }
                    if ("$VAR_RESULT" -ceq "0") {  # Check for errors
                        Write-Log "[MESSAGE] "
                    } else {
                    }
                } else {
                    if ("$INI_SECTION" -like "PA Documents Viewer*") {  # PA Documents Viewer quirk - only copy exes
                        Copy-FileEx -Source "$BUILD_TEMP_PATH\setup\*.exe" -Destination "$TEMP_VAR"  # No overwrite copy files to \\%VAULT_SERVER_ADDRESS%\forqa\...
                    } else {
                        Copy-FileEx -Source "$BUILD_TEMP_PATH\setup\*.*" -Destination "$TEMP_VAR"  # No overwrite copy files to \\%VAULT_SERVER_ADDRESS%\forqa\...
                    }
                    if ("$VAR_RESULT" -ceq "0") {  # Check for errors
                        Write-Log "[MESSAGE] "
                    } else {
                    }
                }
            }
        }
        $LAST_BUILD_VERSION = "$TEMP_VAR\v$BUILD_VERSION"  # LAST_BUILD_VERSION
        #endregion Move all setups to the release folder


        #region Update last build date in ini file
        Write-Log "--- Update last build date in ini file ---"
        # Script block (DelphiScript): update LAST_BUILD_DATE_TIME
        $LAST_BUILD_DATE_TIME = (Get-Date).ToString()
        Set-IniValue -Path "$BUILD_INI" -Section "$INI_SECTION" -Key "LAST_BUILD_DATE_TIME" -Value "$LAST_BUILD_DATE_TIME"  # Set LAST_BUILD_DATE_TIME
        Set-IniValue -Path "$BUILD_INI" -Section "$INI_SECTION" -Key "LAST_BUILD_VERSION" -Value "$LAST_BUILD_VERSION"  # Set LAST_BUILD_VERSION
        #endregion Update last build date in ini file


        #region Commit Source Control changes
        Write-Log "--- Commit Source Control changes ---"
        if ("$SOURCE_CONTROL_LABEL" -ceq "") {  # If SOURCE_CONTROL_LABEL not set
            # [DISABLED] Vault Undo Check Out (Undo check out (FOR TESTING!)) - Undo check out (FOR TESTING!)
            $LABEL = "Build $BUILD_VERSION"  # Set LABEL
            Invoke-VaultCheckIn -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/*" -Comment "$LABEL"  # 
            $LABEL = "$INI_SECTION - build $BUILD_VERSION - $LAST_BUILD_DATE_TIME"  # Set LABEL
            Invoke-VaultLabel -Repository "SDG" -Path "$PROJECT_LABEL_PATH" -Label "$LABEL"  # PROJECT_LABEL_PATH
            # Script block (DelphiScript): generate framework path
            $FRAMEWORK_PATH = ""
            if ($PROJECT_TITLE -like "BANK*") {
                $FRAMEWORK_PATH = "$PA_VS_FRAMEWORK_PATH"
            }
            Invoke-VaultLabel -Repository "SDG" -Path "$FRAMEWORK_PATH" -Label "$LABEL"  # Label framework
            # [DISABLED] Vault Label (Label framework (PA.Blazor.Components)) - Label framework (PA.Blazor.Components)
        } else {
            Invoke-VaultUndoCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/*"  # (nothing should be checked out here anyway)
        }
        #endregion Commit Source Control changes

        #endregion Move all setups to the release folder and commit Source Control changes


        #region Cleanup
        Write-Log "--- Cleanup ---"
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\builds" -Recurse  # Remove builds folder
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\ActiveX\mykey.*"  # Delete key files
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\ActiveX\readme.txt"  # Delete readme.txt file
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\scripts\*.xml"  # Delete scripts\*.xml files
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\scripts\*.log"  # Delete scripts\*.log files
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\scripts\*.remove"  # Delete scripts\*.remove files
        if ("$PROJECT_TITLE" -like "PA Unit Testing*") {  # PA Unit quirk - do not delete the executables
        } else {
            Remove-ItemSafe -Path "$BUILD_TEMP_PATH\*.exe" -Recurse  # Delete *.exe files
        }
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\*.msi" -Recurse  # Delete *.msi files
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH\*.rsm" -Recurse  # Delete *.rsm files
        if (Test-Path "$BUILD_TEMP_PATH\debug_files") {  # If "debug_files" folder exists
            Rename-Item -Path "$BUILD_TEMP_PATH\debug_files" -NewName "debug" -Force  # Rename the debug_files folder to debug
        }
        if (Test-Path "$BUILD_TEMP_PATH\standard_files") {  # Copy standard exe's into the root folder
            Copy-FileEx -Source "$BUILD_TEMP_PATH\standard_files\*.exe" -Destination "$BUILD_TEMP_PATH" -Force  # 
            if ("$PROJECT_TITLE" -like "PA Unit Testing*") {  # PA Unit quirk - also create the latest PA_UNIT_FILE_NAME
                $TEMP_VAR = Split-Path -Path "$PA_UNIT_FILE_NAME" -Parent  # Folder where PA_UNIT_FILE_NAME resides
                if (Test-Path "$TEMP_VAR") {  # If "debug_files" folder exists
                } else {
                    New-Item -ItemType Directory -Path "$TEMP_VAR" -Force | Out-Null  # 
                }
                Copy-FileEx -Source "$BUILD_TEMP_PATH\standard_files\*.exe" -Destination "$TEMP_VAR" -Force  # 
            }
            Remove-ItemSafe -Path "$BUILD_TEMP_PATH\standard_files" -Recurse  # Remove standard_files folder
        }
        if ("$INI_SECTION" -like "PA Documents Viewer*") {  # PA Documents Viewer quirk - cleanup
            Remove-ItemSafe -Path "$BUILD_TEMP_PATH\debug" -Recurse  # Remove debug folder
            Remove-ItemSafe -Path "$BUILD_TEMP_PATH\setup" -Recurse  # Remove setup folder
            Remove-ItemSafe -Path "$BUILD_TEMP_PATH\*.dll" -Recurse  # Delete *.dlli files
        }
        $LOG_TEXT = Get-BuildLog  # Export log
        Set-Content -Path "$BUILD_TEMP_PATH/BuildLog.txt" -Value "$LOG_TEXT"  # Export log to file
        $VersionFolderPath = "$BUILD_TEMP_PATH\..\$V_MAJOR.$V_MINOR.$V_RELEASE"  # set VersionFolderPath
        # [DISABLED] If Directory Exists (VersionFolderPath) - VersionFolderPath
        # [DISABLED] Else
        Copy-Item -Path "" -Destination "" -Recurse -Force  # Copy everything to the VersionFolderPath folder
        Remove-ItemSafe -Path "$BUILD_TEMP_PATH" -Recurse  # 
        #endregion Cleanup

    } else {

        #region FOR TESTING
        Write-Log "--- FOR TESTING ---"
        $TEMP_VAR_3 = " No PASQL log files found"  # Set TEMP_VAR_3
        $VAR_RESULT = 0
        foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\scripts\*.log" -ErrorAction SilentlyContinue)) {  # Check for any PASQL script log files
            $VAR_RESULT_TEXT = $__file.FullName
            $VAR_RESULT++
            $TEMP_VAR = Get-Content -Path "$VAR_RESULT_TEXT" -Raw  # 
            $TEMP_VAR_2 = "LOG_FILE:$VAR_RESULT_TEXT LOG_FILE_TEXT: $TEMP_VAR"
            $TEMP_VAR_2 = "$TEMP_VAR_2" + "`r`n"  # 
            $TEMP_VAR_2 = "$TEMP_VAR_2" + "`r`n"  # 
            $TEMP_VAR_3 = "$TEMP_VAR_3 $TEMP_VAR_2"
        }
        if ("$VAR_RESULT" -ceq "0") {  # If not found
            $TEMP_VAR_3 = " No PASQL log files found"  # Set TEMP_VAR_3
        }
        $VAR_RESULT_TEXT = Get-BuildLog  # 
        Write-Log "[Email skipped] Nightly build for $INI_SECTION succeeded"  # (for testing)
        #endregion FOR TESTING


        #region Cleanup
        Write-Log "--- Cleanup ---"
        # [DISABLED] Remove Directory (Remove build folder) - Remove build folder
        #endregion Cleanup

    }
    if ("$INI_SECTION" -like "Advanced Inquiry*") {  # If building Advanced Inquiry (also build Archive Inquiry)
        $NEXT_PROJECT_TO_BUILD = "38"  # Build Archive Inquiry
        $SOURCE_CONTROL_LABEL = ""  # reset SOURCE_CONTROL_LABEL
        # GOTO: START_OF_THE_BUILD_PROCESS  (GoTo not supported in PS - restructure as loop)
    }
    #endregion Build process

} catch {  # 

    #region Handle errors
    Write-Log "--- Handle errors ---"
    if ("$VAR_RESULT_TEXT" -ne "Unable to locate project in Source Control") {  # If exception was not caused by invalid Source Control path
        Invoke-VaultUndoCheckOut -Repository "SDG" -Path "$SOURCE_CONTROL_SOURCE_PATH/*"  # Undo check out
    } else {
        $VAR_RESULT_TEXT = "$VAR_RESULT_TEXT, path $SOURCE_CONTROL_SOURCE_PATH"
    }
    if ("$NIGHTLY_BUILD" -eq "TRUE") {  # If nightly build
        # [DISABLED] Send E-mail
        # [DISABLED] Send E-mail ((for testing)) - (for testing)

        #region FOR TESTING
        Write-Log "--- FOR TESTING ---"
        $TEMP_VAR_3 = " No PASQL log files found"  # Set TEMP_VAR_3
        $VAR_RESULT = 0
        foreach ($__file in (Get-ChildItem -Path "$BUILD_TEMP_PATH\scripts\*.log" -ErrorAction SilentlyContinue)) {  # Check for any PASQL script log files
            $VAR_RESULT_TEXT = $__file.FullName
            $VAR_RESULT++
            $TEMP_VAR = Get-Content -Path "$VAR_RESULT_TEXT" -Raw  # 
            $TEMP_VAR_2 = "LOG_FILE:$VAR_RESULT_TEXT LOG_FILE_TEXT: $TEMP_VAR"
            $TEMP_VAR_2 = "$TEMP_VAR_2" + "`r`n"  # 
            $TEMP_VAR_2 = "$TEMP_VAR_2" + "`r`n"  # 
            $TEMP_VAR_3 = "$TEMP_VAR_3 $TEMP_VAR_2"
        }
        if ("$VAR_RESULT" -ceq "0") {  # If not found
            $TEMP_VAR_3 = " No PASQL log files found"  # Set TEMP_VAR_3
        }
        $VAR_RESULT_TEXT = Get-BuildLog  # (for testing)
        # [DISABLED] Send E-mail ((for testing)) - (for testing)
        #endregion FOR TESTING

        # [DISABLED] Group (Cleanup) - Cleanup
    } else {
        if ("$VAR_RESULT_TEXT" -notlike "*Build cancelled by user*") {
            Write-Log "[MESSAGE] "
        }
    }
    #endregion Handle errors

    return  # Stop Macro Execution - 
}
