VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GLayout"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public FormPaddingLeft As Long
Public FormPaddingRight As Long
Public FormPaddingTop As Long
Public FormPaddingBottom As Long

Public PaddingLeft As Long
Public PaddingTop As Long

Const cFormPadding As Long = 120
Const cPadding As Long = 120


Private mContainer As Form
Private mLeft As Long
Private mTop As Long
Private mNextLine As Long
Private mWidth As Long
Private mHeight As Long

Public Sub Init(ByRef ctlAny As Form)


 
 FormPaddingLeft = cFormPadding
 FormPaddingRight = cFormPadding
 FormPaddingTop = cFormPadding
 FormPaddingBottom = cFormPadding

 PaddingLeft = cPadding
 PaddingTop = cPadding
 
 Set mContainer = ctlAny
 mLeft = FormPaddingLeft
 mTop = FormPaddingTop
 mNextLine = mTop
 
 mWidth = ctlAny.Width - FormPaddingLeft - FormPaddingRight
 mHeight = ctlAny.Height - FormPaddingTop - FormPaddingBottom
 
End Sub

Private Function perToDbl(ByVal pStr As String) As Double
    pStr = Left$(pStr, Len(pStr) - 1)
    perToDbl = CDbl(pStr) / 100
End Function

Public Sub PutControl(ByRef ctlAny As Object, ByVal iWidth As String, ByVal iHeight As String, ByVal NewLine As Boolean)
    
    Dim rWidth As Long
    Dim rHeight As Long
    
    If mContainer Is Nothing Then Exit Sub
    

    
    If NewLine Then
        mLeft = FormPaddingLeft
        mTop = mNextLine
    End If
    
    If Right$(iWidth, 1) = "%" Then
        rWidth = perToDbl(iWidth) * mWidth
    ElseIf iWidth = 0 Then
        rWidth = mWidth - mLeft
    Else
        rWidth = iWidth
    End If
    
    If Right$(iHeight, 1) = "%" Then
        rHeight = perToDbl(iHeight) * mHeight
    ElseIf iHeight = 0 Then
        rHeight = mHeight - mTop
    Else
        rHeight = iHeight
    End If
    
    
    ctlAny.Move mLeft, mTop, rWidth, rHeight
  
    mLeft = mLeft + rWidth + PaddingLeft
    mNextLine = mTop + rHeight + PaddingTop
    
End Sub



