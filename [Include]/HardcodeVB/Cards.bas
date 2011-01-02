Attribute VB_Name = "Cards"
Option Explicit

Public Const ordFaces = 0
Public Const ordBacks = 1
Public Const ordInvert = 2

Public Const ordCrossHatch = 53
Public Const ordPlaid = 54
Public Const ordWeave = 55
Public Const ordRobot = 56
Public Const ordRoses = 57
Public Const ordIvyBlack = 58
Public Const ordIvyBlue = 59
Public Const ordFishCyan = 60
Public Const ordFishBlue = 61
Public Const ordShell = 62
Public Const ordCastle = 63
Public Const ordBeach = 64
Public Const ordCardHand = 65
Public Const ordUnused = 66
Public Const ordX = 67
Public Const ordO = 68

Public Const ordClubs = 0
Public Const ordDiamonds = 13
Public Const ordHearts = 26
Public Const ordSpades = 39
    
Declare Function cdtInit Lib "Cards32.Dll" ( _
    dx As Long, dy As Long) As Long
Declare Function cdtDrawExt Lib "Cards32.Dll" (ByVal hDC As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, _
    ByVal ordCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Long
Declare Function cdtDraw Lib "Cards32.Dll" (ByVal hDC As Long, _
    ByVal X As Long, ByVal Y As Long, _
    ByVal iCard As Long, ByVal iDraw As Long, ByVal clr As Long) As Long
Declare Function cdtAnimate Lib "Cards32.Dll" (ByVal hDC As Long, _
    ByVal iCardBack As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal iState As Long) As Long
Declare Function cdtTerm Lib "Cards32.Dll" () As Long
