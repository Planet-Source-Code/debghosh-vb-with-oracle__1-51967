Attribute VB_Name = "modLogon"
Option Explicit
Public db As New ADODB.Connection
Public uid As String
Public pwd As String
Public dBase As String

Sub OracleConnect()
    On Error GoTo logonError
    Dim Conn As String
    Dim drv As String
    uid = Trim$(frmLogon.txtUserId.Text)
    pwd = Trim$(frmLogon.txtPassword.Text)
    dBase = Trim$(frmLogon.txtDatabase.Text)
    Set db = New ADODB.Connection
        With frmLogon
        
            If .txtUserId.Text = "" Then
                MsgBox "Please Enter USER ID", vbExclamation
                .txtUserId.SetFocus
                Exit Sub
            ElseIf .txtPassword.Text = "" Then
                MsgBox "Please Enter Password", vbExclamation
                .txtPassword.SetFocus
                Exit Sub
            End If
            'Connection String
            If .txtDatabase.Text <> "" Then
                Conn = "UID= " & uid & ";PWD=" & pwd & ";DRIVER={Microsoft ODBC For Oracle};" _
                & "SERVER=" & dBase & ";"
            Else
                Conn = "UID= " & uid & ";PWD=" & pwd & ";DRIVER={Microsoft ODBC For Oracle};"
            End If
        End With
        
        Screen.MousePointer = vbHourglass
        'Connect With ORACLE
        With db
            .ConnectionString = Conn
            .CursorLocation = adUseClient
            .Open
        End With
        Screen.MousePointer = vbDefault
logonError:
    If Err.Number <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "If Any Error Occured Please Restart the Prog If Reqd. Error Description:" & Err.Description & "", vbCritical
        
        With frmLogon
            .txtUserId.Text = ""
            .txtPassword.Text = ""
            .txtDatabase.Text = ""
            .txtUserId.SetFocus
        End With
    
    Else
        Screen.MousePointer = vbDefault
        Unload frmLogon
        frmMain.Show
    End If
    
End Sub






