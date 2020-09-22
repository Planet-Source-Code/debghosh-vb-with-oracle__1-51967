VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTblDescription 
   Caption         =   "Table Description"
   ClientHeight    =   6060
   ClientLeft      =   1650
   ClientTop       =   1485
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   500
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4965
      TabIndex        =   4
      Top             =   405
      Width           =   1785
   End
   Begin VB.ComboBox cmbLoad 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   270
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   405
      Width           =   4650
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   7215
      Top             =   -105
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4725
      Left            =   3435
      TabIndex        =   2
      Top             =   795
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   8334
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin RichTextLib.RichTextBox rt 
      Height          =   2235
      Left            =   3450
      TabIndex        =   1
      Top             =   5565
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   3942
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmTblDescription.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   7020
      Left            =   255
      TabIndex        =   0
      Top             =   795
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   12383
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuExit 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "frmTblDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Program Decribe How
' you retrieved Table Description.
' send any comment at debughosh@vsnl.net.
Option Explicit
Dim rs As New ADODB.Recordset
Dim b As New clsCommandBtn
Dim n As Node
Dim l As ListItem
Private Sub cmdLoad_Click()
    On Error GoTo rsError
    tv.Nodes.Clear
    Set rs = db.OpenSchema(adSchemaTables, Array(Empty, "" & cmbLoad.Text & "", Empty, "Table"))
    If rs.RecordCount > 0 Then
        Set n = tv.Nodes.Add(, , "TABLESCHEMA", "" & cmbLoad.Text & "")
        n.Expanded = True
        Do Until rs.EOF
            ' Add Table Name
            Set n = tv.Nodes.Add("TABLESCHEMA", tvwChild, "" & cmbLoad.Text & " " & rs!TABLE_NAME & "", rs!TABLE_NAME)
            rs.MoveNext
        Loop
        'Check Whether rs is Closed Or Open, If Open close rs(Recodeset)
        If rs.State = 1 Then
            rs.Close
        End If
    Else
        Set n = tv.Nodes.Add(, , "TSCHEMA", "" & cmbLoad.Text & "")
        n.Expanded = True
        Set n = tv.Nodes.Add("TSCHEMA", tvwChild, , "No Table")
        If rs.State = 1 Then
            rs.Close
        End If
    End If
rsError:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
        If rs.State = 1 Then
            rs.Close
        End If
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
On Error GoTo rsError
    Set b = New clsCommandBtn
    ' Flat Command Button Control
    b.CommandButtonFlat cmdLoad
    Set b = Nothing
    Set rs = New ADODB.Recordset
    ' Add List View Columns
    With lv.ColumnHeaders
        .Add , , "Sr.", 34
        .Add , , "COLUMN NAME", 167
        .Add , , "FLAGS", 100
        .Add , , "NULL", 120
        .Add , , "DATATYPE", 200
        .Add , , "SIZE", 80
        .Add , , "PRECISION", 80
        .Add , , "SCALE", 80
    End With
    Set rs = db.OpenSchema(adSchemaSchemata, Array(Empty, Empty, Empty))
    ' Retrieve Schema Name
    Do Until rs.EOF
        cmbLoad.AddItem rs!SCHEMA_NAME ' Add Schema Name
        rs.MoveNext
    Loop
    cmbLoad.ListIndex = 0
    rs.Close
rsError:
    If Err.Number <> 0 Then
        MsgBox Err.Description
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rs.State = 1 Then
        rs.Close
    End If
    frmMain.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
    frmMain.Show
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    Screen.MousePointer = vbHourglass
On Error GoTo rsError
    If Node.Key = "" & cmbLoad.Text & " " & Node.Text & "" Then
    Dim c As Integer
    c = 1
    lv.ListItems.Clear
    Set rs = db.OpenSchema(adSchemaColumns, Array(Empty, "" & cmbLoad.Text & "", "" & Node.Text & ""))
    Do Until rs.EOF
        Set l = lv.ListItems.Add(, , " " & c & "")
        l.SubItems(1) = rs!COLUMN_NAME
        l.SubItems(2) = rs!COLUMN_FLAGS
        If rs!IS_NULLABLE = 0 Then
            l.SubItems(3) = "NOT NULL"
        Else
            If rs!IS_NULLABLE = -1 Then
                l.SubItems(3) = "NULL"
            Else
                l.SubItems(3) = "UNKNOWN"
            End If
        End If
        Call SetDataType
        c = c + 1
    rs.MoveNext
    Loop
    Call KeyUsed
    End If
    Screen.MousePointer = vbDefault
rsError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "Error : " & Err.Description & " ", vbCritical
        Exit Sub
    End If
End Sub
Private Sub SetDataType()
    Dim x
    Dim d As Integer
    d = rs!DATA_TYPE
    Select Case d
        Case 5 ' Float
            
            If l.SubItems(2) = 120 Or l.SubItems(2) = 24 Then
                l.SubItems(4) = "FLOAT"
                l.SubItems(6) = rs!NUMERIC_PRECISION
                    If rs!NUMERIC_SCALE <> "" Then
                        l.SubItems(7) = rs!NUMERIC_SCALE
                    Else
                        l.SubItems(7) = "NULL"
                    End If
            Else
                l.SubItems(4) = "UNKNOWN"
            End If
        
        Case 128 ' Raw Or Long Raw
            
            If l.SubItems(2) = 104 Then
                l.SubItems(4) = "RAW Or LONG RAW"
                l.SubItems(5) = rs!CHARACTER_MAXIMUM_LENGTH
            Else
                l.SubItems(4) = "UNKNOWN"
            End If
            
        Case 129 ' Char Or Varchar2 Or Long Raw
            
            If l.SubItems(2) = 120 Or l.SubItems(2) = 24 Then
                l.SubItems(4) = "CHAR"
                l.SubItems(5) = rs!CHARACTER_MAXIMUM_LENGTH
            Else
                If l.SubItems(2) = 104 Or l.SubItems(2) = 8 Then
                    l.SubItems(4) = "VARCHAR2"
                    l.SubItems(5) = rs!CHARACTER_MAXIMUM_LENGTH
            Else
                    If l.SubItems(2) = 232 Then
                        l.SubItems(4) = "LONG"
                        l.SubItems(5) = rs!CHARACTER_MAXIMUM_LENGTH
                    Else
                        l.SubItems(4) = "UNKNOWN"
                    End If
                End If
            End If
        
        Case 131 ' Number
            If l.SubItems(2) = 120 Or l.SubItems(2) = 24 Then
                l.SubItems(4) = "NUMBER"
                l.SubItems(6) = rs!NUMERIC_PRECISION
                    If rs!NUMERIC_SCALE <> "" Then
                        l.SubItems(7) = rs!NUMERIC_SCALE
                    Else
                        l.SubItems(7) = "NULL"
                    End If
            Else
                l.SubItems(4) = "UNKNOWN"
            End If
        
        Case 135 'Date
            If l.SubItems(2) = 120 Or l.SubItems(2) = 24 Then
                l.SubItems(4) = "DATE"
            Else
                l.SubItems(4) = "UNKNOWN"
            End If
            
        Case Else
            l.SubItems(4) = "UNKNOWN"
    End Select
End Sub
Private Sub KeyUsed()
    Screen.MousePointer = vbHourglass
    Dim i
    rt.Text = ""
    rt.SelBold = True
    rt.SelColor = &H8000&
    rt.SelText = " Table Name : - " & tv.SelectedItem.Text & "   " & vbCrLf
    rt.SelText = "--------------------------------------------------" & vbCrLf
    
    rt.SelColor = vbBlack
    rt.SelText = "Tablespace: - "
    Set rs = New ADODB.Recordset
        'Tablespace Name
        rs.Open "Select TABLESPACE_NAME From ALL_ALL_TABLES WHERE OWNER='" & cmbLoad.Text & "' AND TABLE_NAME='" & tv.SelectedItem.Text & "'", db, adOpenDynamic, adLockBatchOptimistic
        rt.SelColor = vbBlue
        rt.SelText = rs!TABLESPACE_NAME & vbCrLf
        rt.SelText = vbCrLf
        If rs.State = 1 Then
            rs.Close
        End If
        
        rt.SelColor = vbBlack
        rt.SelText = "Created Time: - "
    Set rs = New ADODB.Recordset
        ' Table Creation Time
        rs.Open "Select CREATED From ALL_OBJECTS Where OWNER='" & cmbLoad.Text & "' AND OBJECT_NAME='" & tv.SelectedItem.Text & "' AND  OBJECT_TYPE='TABLE'", db, adOpenDynamic, adLockBatchOptimistic
        rt.SelColor = vbBlue
        rt.SelText = rs!CREATED & vbCrLf
        rt.SelText = vbCrLf
        If rs.State = 1 Then
            rs.Close
        End If
        
        rt.SelColor = vbBlack
        rt.SelText = "Primary Key : - "
        ' Retrieve Primary Key
    Set rs = db.OpenSchema(adSchemaPrimaryKeys, Array(Empty, "" & cmbLoad.Text & "", "" & tv.SelectedItem.Text & ""))
        rt.SelColor = vbBlue
        If rs.RecordCount > 1 Then
            Do Until rs.EOF
                rt.SelText = "" & rs!COLUMN_NAME & ","
                rs.MoveNext
            Loop
        Else
            Do Until rs.EOF
                rt.SelText = rs!COLUMN_NAME
                rs.MoveNext
            Loop
        End If
        rt.SelText = vbCrLf
        If rs.State = 1 Then
            rs.Close
        End If
        
        rt.SelColor = vbBlack
        rt.SelText = "Foreign Key Table Name : - "
        ' Foreign Key Table Name
    Set rs = db.OpenSchema(adSchemaForeignKeys, Array(Empty, "" & cmbLoad.Text & "", "" & tv.SelectedItem.Text & "", Empty, "" & cmbLoad.Text & "", Empty))
        rt.SelColor = vbBlue
        If rs.RecordCount > 1 Then
            Do Until rs.EOF
                rt.SelText = "" & rs!FK_TABLE_NAME & ","
                rs.MoveNext
            Loop
        Else
            Do Until rs.EOF
                rt.SelText = rs!FK_TABLE_NAME
                rs.MoveNext
            Loop
        End If
        rt.SelText = vbCrLf
        If rs.State = 1 Then
            rs.Close
        End If
        
        rt.SelColor = vbBlack
        rt.SelText = "Foreign Key  : - "
        'Foreign Key
    Set rs = db.OpenSchema(adSchemaForeignKeys, Array(Empty, "" & cmbLoad.Text & "", "" & tv.SelectedItem.Text & "", Empty, "" & cmbLoad.Text & "", Empty))
        rt.SelColor = vbBlue
        If rs.RecordCount > 1 Then
            Do Until rs.EOF
                rt.SelText = "" & rs!FK_COLUMN_NAME & ","
                rs.MoveNext
            Loop
        Else
            Do Until rs.EOF
                rt.SelText = rs!FK_COLUMN_NAME
                rs.MoveNext
            Loop
        End If
        rt.SelText = vbCrLf
        If rs.State = 1 Then
            rs.Close
        End If
        
        rt.SelColor = vbBlack
        On Error Resume Next
        rt.SelText = "Index Name  : - "
        'Index Name
    Set rs = db.OpenSchema(adSchemaIndexes, Array(Empty, "" & cmbLoad.Text & "", Empty, Empty, "" & tv.SelectedItem.Text & ""))
        rt.SelColor = vbBlue
        If rs.RecordCount > 1 Then
            Do Until rs.EOF
                rt.SelText = "" & rs!INDEX_NAME & ","
                rs.MoveNext
            Loop
        Else
            Do Until rs.EOF
                rt.SelText = rs!INDEX_NAME
                rs.MoveNext
            Loop
        End If
        If rs.State = 1 Then
            rs.Close
        End If
        
    Screen.MousePointer = vbDefault
End Sub


