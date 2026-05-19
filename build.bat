call "C:\Builds\Under Development\Products\Collect\Branches\23.1.1\source\rsvars.bat"
rem build release
msbuild "CollectUpgrade.dproj" /verbosity:diag /clp:ShowCommandLine /t:Rebuild /p:Config=Release;platform=Win32;EnvOptionsWarn=false;DCC_ExeOutput=..\;DCC_LocalDebugSymbols=false;DCC_SymbolReferenceInfo=0;DCC_DebugInformation=2 > "C:\Builds\Under Development\Products\Collect\Branches\23.1.1\source\BuildCollectUpgrade.dproj.log"
