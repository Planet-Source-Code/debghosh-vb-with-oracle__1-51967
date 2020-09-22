VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'API For Progress bar
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Private Const WM_USER = &H400
Private Const CCM_FIRST       As Long = &H2000&
Private Const PBM_SETBARCOLOR As Long = (WM_USER + 9)
Private Const CCM_SETBKCOLOR  As Long = (CCM_FIRST + 1)
Private Const PBM_SETBKCOLOR  As Long = CCM_SETBKCOLOR

Public Enum PrScrolling
    Standard = 0
    Smoth = 1
End Enum
' Bar Color
Public Enum PrBarColor
    DRed = vbRed
    DGreen = vbGreen
    DBlue = vbBlue
End Enum
'Appearance Of Progress Bar
Public Enum PrAppearance
    ccFlat = 0
    cc3D = 1
End Enum
Public Sub DProgressBar(pb As ProgressBar, Appr As PrAppearance, BarColor As PrBarColor, DScrolling As PrScrolling)
    With pb
        .Appearance = Appr
        .Scrolling = DScrolling
    End With
    SendMessage pb.hwnd, PBM_SETBARCOLOR, 0, BarColor
    SendMessage pb.hwnd, PBM_SETBKCOLOR, 0, vbWhite
End Sub

