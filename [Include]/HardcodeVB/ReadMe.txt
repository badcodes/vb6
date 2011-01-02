                      Hardcore Visual Basic
               Online Update for Visual Basic 6.0

These zip files contain the sample programs for the update to Hardcore 
Visual Basic. You have the right to use any of this sample code in your
programs. You do not have the right to republish the code without 
permission from Microsoft and Fawcette Technical Publications. 

Sample Programs
---------------

You can examine sample programs in three ways. Let's say you want to
start with the Fun 'n Games program from Chapter 7. This program
illustrates graphics techniques. You can start by running FunNGame.exe
to see what it does. The setup program copies the sample programs to
the Exes directory on your hard disk.

Sample EXE programs will work from your local disk if you registered 
the components using the batch files ToDebug.bat or ToRelease.bat.
FunNGame.exe uses four of the components described in the book. The 
VBCore.dll component provides miscellaneous services. VisualCore.dll 
provides forms wrapped in public classes--Fun 'N Games uses CColorPicker 
and COpenPictureFile. SubTimer.DLL is a small component that provides 
timer and subclassing services--Fun 'n Games uses CTimer. Finally, the 
program uses the XPictureGlass control from PictureGlass.ocx.

After you've clicked on a few buttons and been amazed by swirling
cars, a spiraling ball, and shuffled cards, you'll probably want to
see what makes these tricks work. The next step is to go to the
directory created by the setup program (Hardcore2 by default) and load
the project FunNGame.vbp into the Visual Basic IDE. You can step
through the program and examine any of its code as an example of how
to call the same components from your own programs.

You can stop right there if you want, but most hardcore programmers
will probably want to take the next step and look inside the
components. You can't do that from the VBP file because it references
the compiled components. So you load the project group FunNGame.vbg.
This one includes VBCore.vbp, VisualCore.vbp, and PictureGlass.vbp
(but not SubTimer.vbp). You can step from the program into the more
than 80 classes and modules of VBCore, or check out the five lines of
new code I wrote to create the XPictureGlass control.

If you're really hardcore and want to see the whole thing, you can
load FunNGameDeb.vbg, which adds SubTimer.vbp to the mix. FunNGame.vbg
references the compiled SubTimer.dll and not the project because, as
Chapter 6 explains, it's notoriously difficult to debug subclassing or
timer code that uses callbacks. That's why the code is in the SubTimer
component rather than in VBCore.

Chapter 5 describes the project layout is explained in more detail.

Zip Files and Directory Layout
------------------------------

The following zip files are available for download:

HardCore3.zip	Source for sample program and components plus compare 
                  files, image files, and the Windows API Type Library--
                  everything you need to build the VB6 components and 
                  samples

HardCore35.zip	Same as above, but for VB5

WinTlb3.zip 	Source files for the Windows API Type Library--download 
                  only if you’re interested in Interface Description 
                  Language (IDL)

CppForVB.zip	Some C++ articles included on the book CD--download only
                  if you’re interested in C++ for Visual Basic

WinTlbU.zip	      Unicode version of the Windows API Type Library--slightly 
                  more efficient for programs that will run only on Windows NT

ComponentD.zip	Debug versions of the built VB6 components

ComponentR.zip	Release versions of the built VB6 components

Exes.zip	      Sample programs as VB6 EXE files

ComponentD5.zip	Debug versions of the built VB5 components

ComponentR5.zip	Release versions of the built VB5 components

Exes5.zip	      Sample programs as VB5 EXE files

The Component and Exe files are for your convenience only. You could 
build the same files from the source. The only file you need to get
started is Harcore3.zip or Hardcore35.zip. 

Be sure to use an unzip utility that understands long file names and 
unzips to directories. Nothing will work if the files aren't created 
with their original long names in the correct directories. 

The directory structure created if you unzip all these files looks like
this:

Hardcore3               Sample program project and source file
    Components          Source for components
    Controls            Source for controls
    LocalModules        Private versions of classes and standard modules
    Compare             Compatibility compare files
    Compare             Compatibility files for components
    Release             Release versions of components and controls
    Debug               Debug versions of components and controls
    Exes                Executable sample programs
    TlbSource           Source for the Windows API Type Library  
    CPP		      C++ articles
      Cpp4VB		Four articles in HTML format on writing C++ DLLs
      Sieve			First VBPJ article on writing ATL servers
      ShortCutSvr		Second VBPJ article on writing ATL servers
    
Chapters 1 and 5 describes the differences between Debug and Release 
versions of components.

After unzipping the files, you should run RegisterForVB.bat to register
the Windows API Type Library (used by all the samples) and to copy 
several DLLs used by the samples. This and several other batch files 
require that RegSvr32.exe be in your path. If it isn't, move from its
VB location or edit the batch files to point to the correct location.

RegisterForVB.bat copies Cards32.dll (used in the Fun 'n Games program) 
to your Windows directory. You can delete this DLL if you have no
interest in programming card games. 

It also copies PSAPI.DLL to your Windows directory. This DLL is required 
for iterating through processes and modules under Windows NT. It is described 
in Chapter 6. The WinWatch utility uses it under Windows NT only. If you are 
using Windows 95 or Windows 98, you can delete this file, although you should
save a copy to distribute to customers if you plan to use process iteration 
techniques. 

To uninstall the components, run the UnregisterAll.bat file. There are 
also batch files to register either debug or release versions of the 
components--ToDebug.bat and ToRelease.bat. You can unregister the type 
libraries with RegTlb using the /U option. After unregistering all tools, 
you can delete any directories you on't want.

Bruce McKinney
www.pobox.com/HardcoreVB
