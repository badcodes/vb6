@Echo on

If %1.==d. Goto GotBuild
If %1.==D. Goto GotBuild
If %1.==r. Goto GotBuild
If %1.==R. Goto GotBuild
Echo Syntax: BuildAll <d or r> [, <part>]
Echo.
Echo    d    - Make debug versions
Echo    r    - Make release versions
Echo    part - Optional part: Components, Controls, Sieve, or Clients 
Echo.
Echo You must supply one of the options, r or d, on the command line
Goto Exit2
:GotBuild

If Not %VBEXE%.==. Goto GotVBEXE
Rem Set  VBEXE="C:\Program Files\DevStudio\VB\VB5.EXE"
REM Set  VBEXE="C:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
Set  VBEXE="X:\program\VisualStudio6\VB98\VB6.exe"
:GotVBEXE

Set EXEDIR=Exes
Set CMPDIR=Components
Set CTLDIR=Controls
Set SIVDIR=Sieve
Set LOCDIR=LocalModule
Set VBVER=6
Rem WinNT Start requires "Title"; Win95/98 requires no "Title"!
Set T="BuildAll"
If %PROCESSOR_LEVEL%.==. Set T=

If Exist %VBEXE% Goto ExistVB
Echo This batch file can't find your VB executable at the default location:
Echo.
Echo   %VBEXE%
Echo.
Echo You must supply the executable in the VBEXE environment variable.
Echo For example, use the Set command in a DOS box like this: 
Echo.
Rem Echo   Set VBEXE="D:\Program Files\DevStudio\VB\VB5.EXE"
Echo   Set VBEXE="D:\Program Files\Microsoft Visual Studio\VB98\VB6.EXE"
Echo.
Echo You must include the path in quote marks if it includes spaces. 
Goto Exit
:ExistVB

If %1.==r. Goto DoRelease
If %1.==R. Goto DoRelease
Set BUILD=Debug
Set AF=5
Set COMPDIR=Debug
Goto DidDebug
:DoRelease
Set BUILD=Release
Set AF=0
Set COMPDIR=Release
:DidDebug

Rem Enable this line for testing
Rem Set VBEXE=Echo VBEXE

If Exist BuildAll.log Del BuildAll.log > nul

If %2.==. Goto Components
Goto %2

:Components
Echo.
Echo Building Components...
Echo.
:SubTimer
:subtimer
Echo Building SubTimer...
Start %T% /w %VBEXE% /m %CMPDIR%\SubTimer.vbp /out BuildAll.log /outdir %COMPDIR% /d %CORECONST%
If ErrorLevel 1 Echo Build failed
If %2.==SubTimer. Goto Exit

:VBCore
Echo Building VBCore...
Start %T% /w %VBEXE% /m %CMPDIR%\VBCore.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=1:iVBVer=%VBVER%
If Not ErrorLevel 1 Goto DidVBCore
Echo Build of VBCore failed. Since most other projects depend on VBCore, 
Echo this batch file will terminate. Solve the VBCore problem and try again.
Goto Exit
:DidVBCore
If %2.==VBCore. Goto Exit

:VisualCore
Echo Building VisualCore...
Start %T% /w %VBEXE% /m %CMPDIR%\VisualCore.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=1:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==VisualCore. Goto Exit

:Notify
Echo Building Notify...						
Start %T% /w %VBEXE% /m %CMPDIR%\Notify.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==Notify. Goto Exit
If %2.==Components. Goto Exit

:Controls
Echo.
Echo Building Controls...
Echo.
:ColorPicker
Echo Building ColorPicker...
Start %T% /w %VBEXE% /m %CTLDIR%\ColorPicker.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==ColorPicker. Goto Exit

:DropStack
Echo Building DropStack...
Start %T% /w %VBEXE% /m %CTLDIR%\DropStack.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==DropStack. Goto Exit

:Editor
Echo Building Editor...
Start %T% /w %VBEXE% /m %CTLDIR%\Editor.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==Editor. Goto Exit

:ListBoxPlus
Echo Building ListBoxPlus...
Start %T% /w %VBEXE% /m %CTLDIR%\ListBoxPlus.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==ListBoxPlus. Goto Exit

:PictureGlass
Echo Building PictureGlass...
Start %T% /w %VBEXE% /m %CTLDIR%\PictureGlass.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==PictureGlass. Goto Exit
If %2.==Controls. Goto Exit

:Sieve
Echo.
Echo Building Sieve components, controls, and client...
Echo.
:SieveBasCtlN
Echo Building SieveBasCtlN...
Start %T% /w %VBEXE% /m %SIVDIR%\SieveBasCtlN.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==SieveBasCtlN. Goto Exit

:SieveBasCtlP
Echo Building SieveBasCtlP...
Start %T% /w %VBEXE% /m %SIVDIR%\SieveBasCtlP.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==SieveBasCtlP. Goto Exit

:SieveBasDllN
Echo Building SieveBasDllN...
Start %T% /w %VBEXE% /m %SIVDIR%\SieveBasDllN.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==SieveBasDllN. Goto Exit

:SieveBasDllP
Echo Building SieveBasDllP...
Start %T% /w %VBEXE% /m %SIVDIR%\SieveBasDllP.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==SieveBasDllP. Goto Exit

:SieveBasExeN
Echo Building SieveBasExeN...
Start %T% /w %VBEXE% /m %SIVDIR%\SieveBasExeN.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==SieveBasExeN. Goto Exit

:SieveBasExeP
Echo Building SieveBasExeP...
Start %T% /w %VBEXE% /m %SIVDIR%\SieveBasExeP.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==SieveBasExeP. Goto Exit

:SieveBasGlobalN
Echo Building SieveBasGlobalN...
Start %T% /w %VBEXE% /m %SIVDIR%\SieveBasGlobalN.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==SieveBasGlobalN. Goto Exit

:SieveBasGlobalP
Echo Building SieveBasGlobalP...
Start %T% /w %VBEXE% /m %SIVDIR%\SieveBasGlobalP.vbp /out BuildAll.log /outdir %COMPDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==SieveBasGlobalP. Goto Exit

Rem Attempt to unregister old SieveBasGlobal servers
Rem Some of these will fail, and perhaps all of them, so do silently
RegSvr32 /s C:\Hardcore2\Debug\SieveBasGlobalP.dll
RegSvr32 /s C:\Hardcore2\Release\SieveBasGlobalN.dll
RegSvr32 /s C:\Hardcore2\Debug\SieveBasGlobalP.dll
RegSvr32 /s C:\Hardcore2\Release\SieveBasGlobalN.dll

Echo Registering SieveATL
RegSvr32 /s Release\SieveATL.dll
If ErrorLevel 1 Goto RegisterFailed

:SieveCli
Echo Building SieveCli...
Start %T% /w %VBEXE% /m SieveCli.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
Goto Clients
:RegisterFailed
Echo Registration failed; RegSvr32.exe must be in your path

:Clients
If %2.==SieveCli. Goto Exit
If %2.==Sieve. Goto Exit

Echo.
Echo Building clients...
Echo.
:Addromatic
Echo Building Addromatic...
Start %T% /w %VBEXE% /m Addromatic.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==Addromatic. Goto Exit

:AllAbout
Echo Building AllAbout...
Start %T% /w %VBEXE% /m AllAbout.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==AllAbout. Goto Exit

:AppPath
Echo Building AppPath...
Start %T% /w %VBEXE% /m AppPath.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==AppPath. Goto Exit

:BitBlast
Echo Building BitBlast...
Start %T% /w %VBEXE% /make BitBlast.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==BitBlast. Goto Exit

:Browse
Echo Building Browse...
Start %T% /w %VBEXE% /m Browse.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==Browse. Goto Exit

:BugWiz
Echo Building BugWiz...
Start %T% /w %VBEXE% /m BugWiz.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==BugWiz. Goto Exit

:CollWiz
Echo Building CollWiz...
Start %T% /w %VBEXE% /m CollWiz.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==CollWiz. Goto Exit

:Edwina
Echo Building Edwina...
Start %T% /w %VBEXE% /m Edwina.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==Edwina. Goto Exit

:ErrMsg
Echo Building ErrMsg...
Start %T% /w %VBEXE% /m ErrMsg.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==ErrMsg. Goto Exit

:FunNGame
Echo Building FunNGame...
Start %T% /w %VBEXE% /m FunNGame.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==FunNGame. Goto Exit

:GlobWiz
Echo Building GlobWiz...
Start %T% /w %VBEXE% /m GlobWiz.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==GlobWiz. Goto Exit

:Hardcore
Echo Building Hardcore...
Start %T% /w %VBEXE% /m Hardcore.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==Hardcore. Goto Exit

:LotAbout
Echo Building LotAbout...
Start %T% /w %VBEXE% /m LotAbout.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==LotAbout. Goto Exit

:Meriwether
Echo Building Meriwether...
Start %T% /w %VBEXE% /m Meriwether.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==Meriwether. Goto Exit

:RegTlb
Echo Building RegTlb...
Start %T% /w %VBEXE% /m RegTlb.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==RegTlb. Goto Exit

:RegTlbOld
Echo Building RegTlbOld...
Start %T% /w %VBEXE% /m RegTlbOld.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==RegTlbOld. Goto Exit

:TBezier
Echo Building TBezier...
Start %T% /w %VBEXE% /m TBezier.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TBezier. Goto Exit

:TCollect
Echo Building TCollect...
Start %T% /w %VBEXE% /m TCollect.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TCollect. Goto Exit

:TColorPick
Echo Building TColorPick...
Start %T% /w %VBEXE% /m TColorPick.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TColorPick. Goto Exit

:TCompletion
Echo Building TCompletion...
Start %T% /w %VBEXE% /m TCompletion.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TCompletion. Goto Exit

:TDictionary
Echo Building TDictionary...
Start %T% /w %VBEXE% /m TDictionary.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TDictionary. Goto Exit

:TEdge
Echo Building TEdge...
Start %T% /w %VBEXE% /m TEdge.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TEdge. Goto Exit

:TEnum
Echo Building TEnum...
Start %T% /w %VBEXE% /m TEnum.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TEnum. Goto Exit

:TExecute
Echo Building TExecute...
Start %T% /w %VBEXE% /m TExecute.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TExecute. Goto Exit

:TFolder
Echo Building TFolder...
Start %T% /w %VBEXE% /m TFolder.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TFolder. Goto Exit

:TIcon
Echo Building TIcon...
Start %T% /w %VBEXE% /m TIcon.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TIcon. Goto Exit

:TImage
Echo Building TImage...
Start %T% /w %VBEXE% /m TImage.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TImage. Goto Exit

:TimeIt
Echo Building TimeIt...
Start %T% /w %VBEXE% /m TimeIt.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TimeIt. Goto Exit

:TMessage
Echo Building TMessage...
Start %T% /w %VBEXE% /m TMessage.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TMessage. Goto Exit

:TPalette
Echo Building TPalette...
Start %T% /w %VBEXE% /m TPalette.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TPalette. Goto Exit

:TParse
Echo Building TParse...
Start %T% /w %VBEXE% /m TParse.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TParse. Goto Exit

:TPaths
Echo Building TPaths...
Start %T% /w %VBEXE% /m TPaths.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TPaths. Goto Exit

:TReg
Echo Building TReg...
Start %T% /w %VBEXE% /m TReg.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TReg. Goto Exit

:TRes
Echo Building TRes...
Start %T% /w %VBEXE% /m TRes.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TRes. Goto Exit

:TRes2
Echo Building TRes2...
Start %T% /w %VBEXE% /m TRes2.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TRes2. Goto Exit

:TShare
Echo Building TShare...
Start %T% /w %VBEXE% /m TShare.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TShare. Goto Exit

:TShortcut
Echo Building TShortcut...
Start %T% /w %VBEXE% /m TShortcut.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TShortcut. Goto Exit

:TSort
Echo Building TSort...
Start %T% /w %VBEXE% /m TSort.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TSort. Goto Exit

:TSplit
Echo Building TSplit...
Start %T% /w %VBEXE% /m TSplit.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TSplit. Goto Exit

:TSplit2
Echo Building TSplit2...
Start %T% /w %VBEXE% /m TSplit2.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TSplit2. Goto Exit

:TString
Echo Building TString...
Start %T% /w %VBEXE% /m TString.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TString. Goto Exit

:TSysMenu
Echo Building TSysMenu...
Start %T% /w %VBEXE% /m TSysMenu.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TSysMenu. Goto Exit

:TThread
Echo Building TThread...
Start %T% /w %VBEXE% /m TThread.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TThread. Goto Exit

:TTimer
Echo Building TTimer...
Start %T% /w %VBEXE% /m TTimer.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TTimer. Goto Exit

:TWhiz
Echo Building TWhiz...
Start %T% /w %VBEXE% /m TWhiz.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TWhiz. Goto Exit

:TWindow
Echo Building TWindow...
Start %T% /w %VBEXE% /m TWindow.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==TWindow. Goto Exit

:VB6ToVB5
Echo Building VB6ToVB5...
Start %T% /w %VBEXE% /m VB6ToVB5.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed
If %2.==VB6ToVB5. Goto Exit

:WinWatch
Echo Building WinWatch...
Start %T% /w %VBEXE% /m WinWatch.vbp /out BuildAll.log /outdir %EXEDIR% /d afDebug=%AF%:fComponent=0:iVBVer=%VBVER%
If ErrorLevel 1 Echo Build failed

:Exit
Set EXEDIR=
Set CMPDIR=
Set CTLDIR=
Set SIVDIR=
Set LOCDIR=
Set VBVER=
Set T=
Set BUILD=
Set AF=
Set COMPDIR=

:Exit2

