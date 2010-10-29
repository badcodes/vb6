Attribute VB_Name = "modDHTML"
'PutProperty: Store information in a cookie by calling this
'             function.
'             The required inputs are the named Property
'             and the value of the property you would like to store.
'
'             Optional inputs are:
'               expires : specifies a date that defines the valid life time
'                         of the property.  Once the expiration date has been
'                         reached, the property will no longer be stored or given out.

Public Sub PutProperty(objDocument As HTMLDocument, strName As String, vntValue As Variant, Optional Expires As Date)

     objDocument.cookie = strName & "=" & CStr(vntValue) & _
        IIf(CLng(Expires) = 0, "", "; expires=" & Format(CStr(Expires), "ddd, dd-mmm-yy hh:mm:ss") & " GMT") ' & _

End Sub

'GetProperty: Retrieve the value of a property by calling this
'             function.  The required input is the named Property,
'             and the return value of the function is the current value
'             of the property.  If the proeprty cannot be found or has expired,
'             then the return value will be an empty string.
'
Public Function GetProperty(objDocument As HTMLDocument, strName As String) As Variant
    Dim aryCookies() As String
    Dim strCookie As Variant
    On Local Error GoTo NextCookie

    'Split the document cookie object into an array of cookies.
    aryCookies = Split(objDocument.cookie, ";")
    For Each strCookie In aryCookies
        If Trim(VBA.Left(strCookie, InStr(strCookie, "=") - 1)) = Trim(strName) Then
            GetProperty = Trim(Mid(strCookie, InStr(strCookie, "=") + 1))
            Exit Function
        End If
NextCookie:
        Err = 0
    Next strCookie
End Function


