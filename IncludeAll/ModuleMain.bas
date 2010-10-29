Attribute VB_Name = "ModuleMain"
Option Explicit

Const VBModule = "E:\WorkBench\VB\[Code]\"
Const VBClass = "E:\WorkBench\VB\[Class]\"
Sub Main()
    Dim VBPName As String
    Dim fso As New FileSystemObject
    Dim fs As Files
    Dim f As File
    Dim ts As TextStream
    Dim sContent As String
    Dim sBaseName As String
    Dim sExt As String
    VBPName = Command$
    If VBPName = "" Then VBPName = "IncludeAll.vbp"
    VBPName = fso.GetAbsolutePathName(VBPName)
    Set ts = fso.OpenTextFile(VBPName, ForWriting, True)
    ts.WriteLine "Type=OleDll"
    ts.WriteLine "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\WINDOWS\system32\stdole2.tlb#OLE Automation"
    
    Set fs = fso.GetFolder(VBModule).Files
    For Each f In fs
        sExt = LCase$(fso.GetExtensionName(f.Name))
        If sExt = "bas" Then
        sBaseName = LCase$(fso.GetBaseName(f.Path))
        Mid$(sBaseName, 1, 1) = UCase$(Left$(sBaseName, 1))
        If Left$(sBaseName, 1) <> "M" Then sBaseName = "M" & sBaseName
        ts.WriteLine "Module=" & sBaseName & "; ..\[Code]\" & f.Name
        End If
    Next
    Set fs = fso.GetFolder(VBClass).Files
    For Each f In fs
        sExt = LCase$(fso.GetExtensionName(f.Name))
        If sExt = "cls" Then
        sBaseName = LCase$(fso.GetBaseName(f.Path))
        Mid$(sBaseName, 1, 1) = UCase$(Left$(sBaseName, 1))
        If Left$(sBaseName, 1) <> "C" Then sBaseName = "C" & sBaseName
        ts.WriteLine "Class=" & sBaseName & "; ..\[Class]\" & f.Name
        End If
     Next
    
    sContent = "Startup=!(None)!"
    sContent = sContent & vbCrLf & "Command32=!!"
    sContent = sContent & vbCrLf & "Name=!" & fso.GetBaseName(VBPName) & "!"
    sContent = sContent & vbCrLf & "HelpContextID=!0!"
    sContent = sContent & vbCrLf & "CompatibleMode=!1!"
    sContent = sContent & vbCrLf & "MajorVer=1"
    sContent = sContent & vbCrLf & "MinorVer=0"
    sContent = sContent & vbCrLf & "RevisionVer=0"
    sContent = sContent & vbCrLf & "AutoIncrementVer=1"
    sContent = sContent & vbCrLf & "ServerSupportFiles=0"
    sContent = sContent & vbCrLf & "CompilationType=0"
    sContent = sContent & vbCrLf & "OptimizationType=0"
    sContent = sContent & vbCrLf & "FavorPentiumPro(tm)=0"
    sContent = sContent & vbCrLf & "CodeViewDebugInfo=0"
    sContent = sContent & vbCrLf & "NoAliasing=0"
    sContent = sContent & vbCrLf & "BoundsCheck=0"
    sContent = sContent & vbCrLf & "OverflowCheck=0"
    sContent = sContent & vbCrLf & "FlPointCheck=0"
    sContent = sContent & vbCrLf & "FDIVCheck=0"
    sContent = sContent & vbCrLf & "UnroundedFP=0"
    sContent = sContent & vbCrLf & "StartMode=1"
    sContent = sContent & vbCrLf & "Unattended=0"
    sContent = sContent & vbCrLf & "Retained=0"
    sContent = sContent & vbCrLf & "ThreadPerObject=0"
    sContent = sContent & vbCrLf & "MaxNumberOfThreads=1"
    sContent = sContent & vbCrLf & "ThreadingModel=1"
    sContent = sContent & vbCrLf & "[MS Transaction Server]"
    sContent = sContent & vbCrLf & "AutoRefresh = 1"
    sContent = Replace$(sContent, "!", Chr(34))
    
    ts.WriteLine sContent
    ts.Close
    
End Sub
