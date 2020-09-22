VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommandBtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Flat Command Button Control.. I get this Code from Planet-source-code.com. Only I did it by Class.
Option Explicit

Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)

Private bStyle As Long
Private Const BS_3STATE As Long = &H5&
Private Const BS_CHECKBOX As Long = &H2&
Private Const BS_PUSHBUTTON As Long = &H0&
Private Const BS_PUSHLIKE As Long = &H1000&
Private Const BS_PATTERN As Long = 3
Private Const BS_PATTERN8X8 As Long = 7
Private Const GWL_EXSTYLE As Long = -20
Public Sub CommandButtonFlat(cBtn As CommandButton)
    InitialBtnStyle cBtn
    If GetWindowLong&(cBtn.hwnd, GWL_STYLE) And BS_FLAT Then
        Exit Sub
    End If
    SetWindowLong cBtn.hwnd, GWL_STYLE, bStyle Or BS_FLAT
    cBtn.Refresh
End Sub
Private Sub InitialBtnStyle(cBtn As CommandButton)
    bStyle = GetWindowLong(cBtn.hwnd, GWL_STYLE)
End Sub



