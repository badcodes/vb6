@Echo Off

ChDir Debug

Echo Registering DLL components (debug)
For %%I IN (*.DLL) Do RegSvr32 /s %%I
Echo Registering OCX components (debug)
For %%I IN (*.OCX) Do RegSvr32 /s %%I
Echo Registering EXE servers (debug)
For %%I IN (*.EXE) Do %%I /regserver

CHDIR ..

RegSvr32 /s Release\SieveATL.DLL

ECHO Done
