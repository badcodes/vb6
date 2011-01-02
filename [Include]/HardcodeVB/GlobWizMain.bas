Attribute VB_Name = "MGlobWizMain"
Option Explicit

Sub Main()
    If Command$ = sEmpty Then
        Dim frm As New FGlobalWizard
        frm.Show
    Else
        Dim sType As String, sFileSrc As String
        sType = GetToken(Command$, " ")
        sFileSrc = GetToken(sEmpty, " ")
        ' Select the appropriate filter and assign to any old object
        Dim filterobj As Object
        Select Case sType
        Case "/pubpriv"
            ' Translates public class to private class
            Set filterobj = New CPubPrivFilter
        Case "/globmod"
            ' Translates global class to standard module
            Set filterobj = New CGlobModFilter
        End Select
        
        Dim filter As IFilter
        Set filter = filterobj
        filter.Source = GetFileText(sFileSrc)
        FilterText filter
        
        Dim sFileDst As String
        Select Case sType
        Case "/pubpriv"
            sFileDst = "P_" & Right$(filterobj.Name, Len(filterobj.Name) - 1) & ".cls"
        Case "/globmod"
            sFileDst = Right$(filterobj.Name, Len(filterobj.Name) - 1) & ".bas"
        End Select
        
        SaveFileStr sFileDst, filter.Target
    End If
End Sub
