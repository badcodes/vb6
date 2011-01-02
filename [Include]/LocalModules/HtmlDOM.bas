Attribute VB_Name = "MHtmlDom"
Option Explicit

Public Function SetAttribute(ByRef vDoc As HTMLDocument, ByVal vId As String, ByVal vName As String, ByVal vValue As String) As Boolean
    On Error GoTo ErrorSet
    Dim elm As IHTMLElement
    Set elm = vDoc.getElementById(vId)
    elm.SetAttribute vName, vValue
    SetAttribute = True
    Exit Function
ErrorSet:
    Debug.Print Err.Description
    SetAttribute = False
End Function


Public Function FindElement(ByRef vResult As IHTMLElement, ByRef vDoc As HTMLDocument, ByVal vId As String) As Boolean
    Dim elm As IHTMLElement
    Set elm = vDoc.getElementById(vId)
    If elm Is Nothing Then
        Exit Function
    Else
        Set vResult = elm
        FindElement = True
    End If
End Function
