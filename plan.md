# Plan: PAApplications.bxp → PowerShell Build Script

## Problem Statement
Converting a SmartBear BuildStudio .bxp build script (41K lines XML, 1,425 operations, 77 types, 39 projects) into a PowerShell 7 script. The script automates building Delphi applications: source control → config updates → compile → test → package → sign → release.

## Current State
- **Converter** (`_converter.py`) handles all 77 operation types, generates ~2,700-line PS1
- **Infrastructure working**: INI parsing, Vault auth (autobuild/autobuild), WinForms dialogs, project selection, Delphi version detection
- **NOT working yet**: 25 embedded script blocks (DelphiScript/JScript/VBScript) are comments only — these contain critical build logic

## Build Pipeline (8 Phases)

### Phase 1: Read Settings
1. Get current year (DelphiScript #3 → needs conversion)
2. Check PAApplications.ini exists
3. Show project selection dialog (radio group) — WORKING
4. Read INI values: PROJECT_TITLE, SOURCE_CONTROL paths, HELP_FILE_LOCATION
5. Detect Delphi version (INI → path → default 10.4 → user confirm) — WORKING
6. Read SUPPORTS_* flags (Oracle, SUN4/5/6)

### Phase 2: Get Highest Available Project Version from Source Control
1. Vault GetLatest on source path
2. Set working folder for Vault
3. File enumerator on Vault to find version folders
4. Version comparison (DelphiScript #4-5, JScript) — needs conversion
5. Determine SOURCE_CONTROL_VERSION (highest available)

### Phase 3: Update Version Numbers
1. Read current BUILD_VERSION from Vault
2. Version comparison (VBScript #6-8) — needs conversion
3. Show version selection dialog (increment build/release/minor/major)
4. Parse version into V_MAJOR, V_MINOR, V_RELEASE, V_BUILD (VBScript #9)
5. Increment selected component (VBScript #10-13) — needs conversion

### Phase 4: Update Config Files & Prepare Build
1. Vault GetLatest on project files
2. Update .dcc32.cfg / .dcc64.cfg with library paths per Delphi version
3. Update .dproj XML with version info
4. Build compiler command string (DelphiScript #1) — **CRITICAL, needs conversion**
5. Check/update EurekaLog settings (.eof files)
6. Build project icon name (DelphiScript #16)

### Phase 5: Compile Delphi Projects
1. File enumerator: loop all *.dpr files in project
2. For each DPR:
   a. Run submacro "Update dcc config file" — patches library paths
   b. Execute dcc32.exe or dcc64.exe with flags: `-B -E..\ -U"paths" project.dpr`
   c. Run submacro "Check build log file" — parse for Fatal/Error
3. Run PASQL tests (if applicable)
4. Run PA Unit tests (PAUnitCMD.exe) with date comparison (DelphiScript #17)

### Phase 6: Build Setups (Installers)
1. Construct Wise installer string (DelphiScript #21) — needs conversion
2. Execute WiseInstaller for each setup
3. Sign executables (signfile.ps1 integration)
4. Copy to release folder

### Phase 7: Commit & Release
1. Vault Check In updated version files
2. Vault Label with version string
3. Copy builds to network release location
4. File overwrite prompts if destination exists

### Phase 8: Cross-Project Dependencies
1. If INI_SECTION starts with "Advanced Inquiry":
   - Set NEXT_PROJECT_TO_BUILD = Archive Inquiry index
   - Loop back to Phase 1 (build Archive Inquiry automatically)

### Error Handling (Catch Block)
1. Vault UndoCheckOut on all modified files
2. Log error details
3. Clean up temporary files

---

## 25 Script Blocks Requiring Manual Conversion

### Priority 1 — Build-Critical (blocks the compile step)
| # | Description | Language | Complexity |
|---|------------|----------|-----------|
| 1 | Construct compiler command string (StandardBuild/DebugBuild) | DelphiScript | HIGH — parses DPR, builds dcc32 cmd line |
| 2 | Replace CR with CRLF in text | DelphiScript | LOW — string replace |
| 3 | Get current year | DelphiScript | LOW — `(Get-Date).Year` |

### Priority 2 — Version Logic (blocks version management)
| # | Description | Language | Complexity |
|---|------------|----------|-----------|
| 4-5 | Set SOURCE_CONTROL_VERSION to higher of two | JScript | MEDIUM — version comparison |
| 6-8 | Set BUILD_VERSION to higher of two | VBScript | MEDIUM — version comparison |
| 9 | Parse version into V_MAJOR/MINOR/RELEASE/BUILD | VBScript | LOW — string split |
| 10 | V_BUILD + 1 | VBScript | LOW — increment |
| 11 | V_RELEASE + 1, V_BUILD = 0 | VBScript | LOW — increment + reset |
| 12 | V_MINOR + 1, V_RELEASE = 0, V_BUILD = 0 | VBScript | LOW — increment + reset |
| 13 | V_MAJOR = current year, reset rest | VBScript | LOW — year + reset |

### Priority 3 — Build Support
| # | Description | Language | Complexity |
|---|------------|----------|-----------|
| 16 | Build project icon name | DelphiScript | LOW — string replace |
| 17 | Compare dates (PA Unit exe vs MockDatabaseSettings) | DelphiScript | MEDIUM — date parsing |
| 20 | Set TEMP_VAR to higher version | VBScript | MEDIUM — version compare |
| 21 | Build Wise installer string | DelphiScript | MEDIUM — string construction |

### Priority 4 — Path/Framework Logic
| # | Description | Language | Complexity |
|---|------------|----------|-----------|
| 14-15 | Unnamed scripts | DelphiScript | UNKNOWN — need investigation |
| 18-19 | Unnamed scripts | DelphiScript | UNKNOWN — need investigation |
| 22 | Set SOURCE_CONTROL_TRUNK_PATH | JScript | LOW — path strip |
| 23 | Set SOURCE_CONTROL_VISUAL_STUDIO_PATH | JScript | LOW — path construction |
| 24 | Save build datetime | DelphiScript | LOW |
| 25 | Generate framework path | DelphiScript | MEDIUM |

---

## Implementation Todos

### Group A: Script Block Conversions (Priority 1-2)
- convert-year-script: Convert "Get current year" DelphiScript → `$CURRENT_YEAR = (Get-Date).Year`
- convert-crlf-script: Convert "Replace CR with CRLF" → PS string replace
- convert-build-string: Convert "Construct compiler command" DelphiScript → PS (CRITICAL)
- convert-version-compare: Convert version comparison scripts (4-8, 20) → PS `[version]` type
- convert-version-increment: Convert version parse/increment scripts (9-13) → PS

### Group B: Script Block Conversions (Priority 3-4)
- convert-icon-name: Convert "Build project icon name" → PS string replace
- convert-date-compare: Convert date comparison scripts → PS `[datetime]`
- convert-wise-string: Convert Wise installer string builder → PS
- convert-path-scripts: Convert path/framework scripts (22-25) → PS
- investigate-unnamed: Examine scripts #14-15, #18-19 for actual purpose

### Group C: Integration & Testing
- integrate-signfile: Wire signfile.ps1 into the build pipeline
- fix-vault-paths: Ensure all vault operations use correct server/paths
- test-vault-flow: End-to-end test of Vault get/checkin/label cycle
- test-compile-flow: Test Delphi compilation with correct compiler paths
- add-ini-delphi-versions: Add DELPHI_VERSION to each INI section
- cross-project-loop: Verify Advanced→Archive Inquiry loop works
- remove-debug-logging: Clean up [DEBUG] lines once stable
- update-documentation: Update PAApplications.md with final state

---

## Notes
- Delphi compiler paths in BXP reference `C:\Compilers\Delphi XE2\...` but actual installs are at `C:\Program Files (x86)\Embarcadero\...` — need mapping
- Vault server: sdg1.pa.com.au, credentials: autobuild/autobuild
- Only XE2, XE6, and 10.4 Sydney are installed on this machine (no Delphi 6 or 2007)
- Nightly build mode is not needed
- Email sending is removed (replaced with Write-Log)
- The converter should be kept in sync with manual PS1 edits
