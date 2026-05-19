# PAApplications Build Script

## Overview

This PowerShell script automates the build process for PA Applications products. It handles source control operations, version management, Delphi compilation, setup building, unit testing, and release deployment.

## Origin

This script was converted from a **SmartBear BuildStudio** build script (`PAApplications.bxp`) — a proprietary XML-based build automation format. The original script contained approximately 29,000 lines of XML defining ~1,050 operations across 68 operation types.

The conversion was performed using a Python-based converter (`_converter.py`) that parses the BXP XML structure and generates equivalent PowerShell code targeting **PowerShell 7**.

## Status

> **In progress**
>
> `PAApplications.ps1` has been converted to PowerShell, but the conversion is still being stabilized. The current focus is splitting reusable logic into modules and adding tests so converted script blocks can be verified without running a full build.
>
> Notes:
> - Module extraction has started under `Modules\`.
> - Per-module tests live under `Tests\` and can be run with `.\Invoke-Tests.ps1`.
> - Vault, Delphi compiler, installer, signing, and network-share flows still require environment-specific verification.
> - Encrypted credentials (Vault password) should still be migrated to a secure credential store.

## How to Run

### Prerequisites

- PowerShell 7 or later (`pwsh.exe`)
- Delphi build tools and project sources
- Network access to source control and build output paths

### Launch

Use the included shortcut (`PAApplications.lnk`) or run directly:

```powershell
pwsh -NoExit -ExecutionPolicy Bypass -File PAApplications.ps1
```

On launch the script presents a project selection dialog, then a version increment dialog, and proceeds through the full build pipeline for the selected product.

## Build Pipeline

1. **Project Selection** — interactive menu to choose which product to build
2. **Version Management** — increment build/release/minor version numbers
3. **Source Control** — get latest sources from Vault
4. **Version Stamping** — update version info in Delphi project files and INI configs
5. **Compilation** — build Delphi projects via MSBuild
6. **Unit Testing** — run PA unit tests against test databases
7. **Setup Building** — create installers via InstallAware
8. **Release Deployment** — copy build outputs to release paths
9. **Source Control Commit** — label and commit version changes
10. **Notification** — email build results

## Files

| File | Description |
|------|-------------|
| `PAApplications.ps1` | The generated PowerShell build script |
| `PAApplications.bxp` | Original BuildStudio source (XML) |
| `PAApplications.lnk` | Shortcut to launch the script |
| `PAApplications.ini` | Build configuration (project paths, settings) |
| `_converter.py` | Python converter that generates the PS1 from the BXP |
| `PAApplications.md` | This document |
| `Modules\*.psm1` | Extracted helper modules for config, versioning, files, Delphi build commands, Vault, signing, installers, and test parsing |
| `Tests\*.Tests.ps1` | Pester tests for the extracted modules |
| `Invoke-Tests.ps1` | Test runner for the Pester suite |
