Attribute VB_Name = "modDBManager"
Option Explicit
'Public Const DBConString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Krisana\Projects\AWS_Stock\aws_stock.mdb;Persist Security Info=False"
Public DBConString As String
'Public DBManager As AWSDBManager
Public DBConn As New ADODB.Connection

'Public Const DSN_NAME As String = "DSN_MMS"

Public Function DBConnExc(ByVal tSql As String) As Boolean
On Error GoTo ErrD
Dim i As Integer, j As Integer
Retry:
    On Error GoTo ErrD
    DBConnExc = True
    If GetDBConnection = False Then
        DBConnExc = False
        i = i + 1
        If i <= 3 Then GoTo Retry
        MsgBox "การดำเนินการยังไม่เสร็จสิ้น ไม่สามารถดำเนินการต่อได้ กรุณาลองใหม่", vbCritical
        Exit Function
    End If
    DBConn.Execute tSql
    ClostDBConnection
    Exit Function
ErrD:
    j = j + 1
    If j <= 3 Then GoTo Retry
    DBConnExc = False
    MsgBox "การดำเนินการยังไม่เสร็จสิ้น ไม่สามารถดำเนินการต่อได้ กรุณาลองใหม่", vbCritical
End Function

Public Function GetDBConnection(Optional conStr As String) As Boolean
On Error GoTo GetDBConnectionErr:
Dim conStr_ As String, i As Integer
    GetDBConnection = True
    If DBConn.State = adStateOpen Then DBConn.Close
    If conStr <> "" Then
        conStr_ = conStr
    Else
        conStr_ = DBConString
    End If
    
RetryCon:
    i = i + 1
    Set DBConn = New ADODB.Connection
    DBConn.ConnectionTimeout = 60
    DBConn.CommandTimeout = 180
    DBConn.CursorLocation = adUseServer
    DBConn.Open conStr_
    Exit Function
    
GetDBConnectionErr:
    If i > 3 Then
        GetDBConnection = False
        MsgBox "ไม่สามารถติดต่อ Server ได้"
        Exit Function
    End If
    If Err.Number = "-2147467259" Then GoTo RetryCon
    Exit Function
End Function

Public Sub ClostDBConnection()
On Error Resume Next
    DBConn.Close
    Set DBConn = Nothing
End Sub

Public Function GetRS(ByVal tSql As String, ByVal CursorType_ As CursorTypeEnum, ByVal LockType_ As LockTypeEnum) As ADODB.Recordset
On Error GoTo ErrD
Dim RS As New ADODB.Recordset
Retry:
    GetDBConnection
    With RS
        .ActiveConnection = DBConn
        .CursorLocation = adUseClient
        .LockType = LockType_
        .CursorType = CursorType_
        .Source = tSql
        .Open
        Set .ActiveConnection = Nothing
    End With
    Set GetRS = RS
    Set RS = Nothing
    ClostDBConnection
Exit Function
ErrD:
    If Err.Number = -2147467259 Then
        DoEvents
        MsgBox "การเชื่อมต่อกับเซิฟเวอร์ถูกตัดขาด" & vbCrLf & "กรุณาลองเปิดโปรแกรมใหม่"
        Exit Function
    End If
'    MsgBox Err.Description, vbCritical, "Error"
    Set RS = Nothing
End Function
