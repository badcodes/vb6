@ECHO OFF

CHDIR Debug
ECHO Unregistering components and servers (debug)
FOR %%I IN (*.DLL) DO RegSvr32.exe /s /u %%I
FOR %%I IN (*.EXE) DO %%I /unregserver

CHDIR ..\Release
ECHO Unregistering components and servers (release)
FOR %%I IN (*.DLL) DO RegSvr32.exe /s /u %%I
FOR %%I IN (*.EXE) DO %%I /unregserver

CHDIR ..

ECHO Done
