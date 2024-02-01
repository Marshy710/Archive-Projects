Attribute VB_Name = "modMain"
Option Explicit

Global DBConn                       As ADODB.Connection
Global db_name                      As String
Global db_server                    As String
Global db_port                      As String
Global db_user                      As String
Global db_pass                      As String
Global varStr                       As String
Global ConString                    As String
Global varConnection                As Boolean
Global varUserAdmin                 As Boolean
Global varUsersId                   As Integer

Sub Main()
Dim UserLogin As Boolean
    ConDatabaseConnection
    
    With frmLogin2
        .Show vbModal
        UserLogin = .LoginSucceeded
    End With
    
    If UserLogin And varUserAdmin Then
        frmMain.Show
    ElseIf UserLogin Then
        frmMainUser.Show
    End If
End Sub

Public Sub ConDatabaseConnection()
    db_name = "Database11.accdb"
    varStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
    App.Path & "\Database\" & db_name & ";Persist Security Info=False"
    Set DBConn = New ADODB.Connection
    DBConn.Open varStr
    DataEnvironment1.Connection1.Open varStr
End Sub

Public Sub conTable(RecSet As ADODB.Recordset, sqlString As String)
    Set RecSet = New ADODB.Recordset
    RecSet.CursorLocation = adUseClient
    RecSet.Open sqlString, DBConn, adOpenDynamic, adLockOptimistic
End Sub
