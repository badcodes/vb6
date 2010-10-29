@Echo Off

ChDir Release

Echo Registering DLL components (release)
For %%I IN (*.DLL) Do RegSvr32 /s %%I
Echo Registering OCX components (release)
For %%I IN (*.OCX) Do RegSvr32 /s %%I
Echo Registering EXE servers (release)
For %%I IN (*.EXE) Do %%I /regserver

CHDIR ..

ECHO Done
