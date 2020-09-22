VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Logon (Enter USER ID,Password And Database)"
   ClientHeight    =   3900
   ClientLeft      =   2250
   ClientTop       =   2175
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6405
   Begin VB.TextBox txtDatabase 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1500
      TabIndex        =   2
      Top             =   2175
      Width           =   4500
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1500
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1380
      Width           =   4500
   End
   Begin VB.TextBox txtUserId 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1500
      TabIndex        =   0
      Top             =   720
      Width           =   4500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4515
      TabIndex        =   4
      Top             =   3250
      Width           =   1440
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3030
      TabIndex        =   3
      Top             =   3250
      Width           =   1440
   End
   Begin VB.Label Label4 
      Caption         =   "Right Click On any Text Box. Copy Or Paste Any Text On Text Box"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   375
      TabIndex        =   8
      Top             =   360
      Width           =   5760
   End
   Begin VB.Line Line1 
      X1              =   90
      X2              =   6285
      Y1              =   3105
      Y2              =   3105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Database"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   170
      TabIndex        =   7
      Top             =   2220
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   170
      TabIndex        =   6
      Top             =   1485
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "USER ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   170
      TabIndex        =   5
      Top             =   765
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   3615
      Left            =   90
      Shape           =   4  'Rounded Rectangle
      Top             =   195
      Width           =   6210
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim p As New clsCommandBtn
Dim t As New clsTextSubclass

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    OracleConnect
End Sub

Private Sub Form_Load()
    Set p = New clsCommandBtn
    p.CommandButtonFlat cmdOk
    p.CommandButtonFlat cmdCancel
    Set p = Nothing
    Set t = New clsTextSubclass
    t.TxtPop txtUserId
    t.TxtPop txtPassword
    t.TxtPop txtDatabase
End Sub
Private Sub Form_Unload(Cancel As Integer)
    t.txtUnhook
    Set t = Nothing
End Sub
