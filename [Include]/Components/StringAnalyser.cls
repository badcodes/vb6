VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStringAnalyser"
Attribute VB_Creatable = False
Attribute VB_Exposed = True
Option Explicit

'Class to provide lNearestDelimiterPositionethods of popping words from a string
'given a set of pcolDelimiters
'The pcolDelimiters lNearestDelimiterPositionay be variable length as as lNearestDelimiterPositionany as you like
'There is no precedence of pcolDelimiters
'The first one in the string is the first delimiter chosen

Private bStillInitializing As Boolean
Private sCopyOfInputString As String
Private colOfDelimiters As New Collection
Private bIgnoreCase As Boolean
Private lLastDelimiterPosition As Long
Private lLatestDelimiterPosition As Long


Private Sub Class_Initialize()

On Error GoTo Class_InitializeERR

    bStillInitializing = True
    lLastDelimiterPosition = 1
    lLatestDelimiterPosition = 1
    
Class_InitializeEXIT:
    Exit Sub
Class_InitializeERR:
    Err.Raise (Err.Number)
    Resume Class_InitializeEXIT
End Sub



Public Function PopWord(Optional pInputString As Variant, _
                        Optional pbIgnoreCase As Variant, _
                        Optional pcolDelimiters As Variant) As Variant

'Parameters
'   pInputString = string to be analysed
'   pbIgnoreCase = flag to indicate if case should be considered
'                  when searching for pcolDelimiters
'   pcolDelimiters = collection of delimiter values
'Result
'   NULL if no more delimited words found
'   <VALUE> if a word is found
'Note
'   If the result is not NULL, there may be an empty string
'   returned.  This occurs when you have two or delimiters next
'   to each other

On Error GoTo popWordERR

Dim sDelimiter As Variant
Dim lNearestDelimiterPosition As Long
Dim lLengthOfNearestDelimiter As Long

    'Because the constructor cannot accept arguments
    'We must finish initialisation here
    'The first call to pop word is with parameters
    'Subsequent calls do not require parameters
    'and will be ignored if passed
    If bStillInitializing Then
        bStillInitializing = False
        
        'save the string to be analysed
        sCopyOfInputString = pInputString
        
        If IsMissing(pbIgnoreCase) Then
            'user wants default for ignore case
            bIgnoreCase = True
        Else
            'otheriwse use user supplied value
            bIgnoreCase = pbIgnoreCase
        End If
            
        If IsMissing(pcolDelimiters) Then
            'use our default set of pcolDelimiters
            colOfDelimiters.Add " "
        Else
            'otherwise use user supplied list
            Set colOfDelimiters = pcolDelimiters
        End If
    End If
    
    'we have reached the end of the string to be analysed
    'return NULL
    If lLatestDelimiterPosition >= Len(sCopyOfInputString) Then
        PopWord = Null
        Exit Function
    End If
    
    'default dummy value
    lNearestDelimiterPosition = 99999999
    
    'For each of the delimiters
    For Each sDelimiter In colOfDelimiters
        'change comparison depending on case parameter
        If bIgnoreCase Then
            lLatestDelimiterPosition = InStr(lLastDelimiterPosition, UCase(sCopyOfInputString), UCase(sDelimiter))
        Else
            lLatestDelimiterPosition = InStr(lLastDelimiterPosition, sCopyOfInputString, sDelimiter)
        End If
        
        'this delimiter came before any other, save position
        If (lLatestDelimiterPosition <= lNearestDelimiterPosition) And (lLatestDelimiterPosition <> 0) Then
            lNearestDelimiterPosition = lLatestDelimiterPosition
            lLengthOfNearestDelimiter = Len(sDelimiter)
        End If
    Next
    
    'we didn't allocate anything to lNearestDelimiterPositio must be no colOfDelimiters
    'remaining in the input string
    If lNearestDelimiterPosition = 99999999 Then
        lLatestDelimiterPosition = Len(sCopyOfInputString) + 1
        lNearestDelimiterPosition = lLatestDelimiterPosition
    End If
    
    If lNearestDelimiterPosition <> 99999999 Then
        PopWord = Mid(sCopyOfInputString, lLastDelimiterPosition, (lNearestDelimiterPosition - lLastDelimiterPosition))
        lLastDelimiterPosition = lNearestDelimiterPosition + lLengthOfNearestDelimiter
    End If
    
popWordEXIT:
    Exit Function
popWordERR:
    PopWord = Null
    lLatestDelimiterPosition = Len(sCopyOfInputString) + 1
    Err.Raise Err.Number
    Resume popWordEXIT
End Function


Public Sub Restart()

'Reinitialises the class so the user can start serching for another string
'User could set the class to nothing and declare another instance

On Error GoTo RestartERR

    Class_Initialize
    
RestartEXIT:
    Exit Sub
RestartERR:
    Err.Raise (Err.Number)
    Resume RestartEXIT
End Sub


