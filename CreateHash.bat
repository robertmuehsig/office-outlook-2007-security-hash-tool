@echo off
IF "%1" == "/register" GOTO Install
IF "%1" == "/unregister" GOTO Uninstall
IF "%1" == "" GOTO Usage

cscript //I //NoLogo addinhash.vbs %1
GOTO End

:Usage
echo.
echo Missing COM add-in filename
echo.
echo Installation: CreateHash.bat /install
echo.
echo Usage: CreateHash.bat filename.dll
echo Where filename.dll is the path to the add-in file that should be trusted
echo.
GOTO End

:Install
echo.
echo Registering hashctl.dll...
%windir%\system32\regsvr32.exe /s hashctl.dll
GOTO End

:Uninstall
echo.
echo Unregistering hashctl.dll...
%windir%\system32\regsvr32.exe /s /u hashctl.dll

:End
