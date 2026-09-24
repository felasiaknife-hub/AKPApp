Attribute VB_Name = "modDBManager"
Option Explicit
'Public Const DBConString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Krisana\Projects\AWS_Stock\aws_stock.mdb;Persist Security Info=False"
Public DBConString As String
Public DBConStock As String
Public DBManager As AWSDBManager  '// เปิดสำหรับโปรแกรม Waste Data ปิดสำหรับ Waste Management & Commission
Public DBConn As New ADODB.Connection
Public DBExec As New ADODB.Connection
Public cnMain As New ADODB.Connection
'Public RS As New ADODB.Recordset

Type ParamType
    ParamName As String
    ParamType   As ADODB.DataTypeEnum
    ParamInput As ADODB.ParameterDirectionEnum
    ParamValue As Variant
    ParamValueLen As Variant
End Type
Public Param() As ParamType

Public MoveWaste As String
Public Years_  As String
Public WasteID As String
Public Wastelab_ID As String
Public countID As String

Public Function DBConnExc(ByVal tSql As String) As Boolean
    DBConnExc = True
    If GetDBExecConnec = False Then
        DBConnExc = False
        MsgBox "การดำเนินการยังไม่เสร็จสิ้น ไม่สามารถดำเนินการต่อได้ กรุณาลองใหม่", vbCritical
        Exit Function
    End If
    DBExec.Execute tSql
    CloseDBExec
End Function

Private Function GetDBExecConnec(Optional conStr As String) As Boolean
On Error GoTo GetDBConnectionErr:
Dim conStr_ As String, i As Integer
    If GetDBExecConnec = True Then Exit Function
    If DBExec.State = adStateOpen Then DBExec.Close
    If conStr <> "" Then
        conStr_ = conStr
    Else
        conStr_ = DBConString
    End If
    
RetryCon:
On Error GoTo GetDBConnectionErr:
    i = i + 1
    GetCnString
    Set DBExec = New ADODB.Connection
    DBExec.ConnectionTimeout = 15
    DBExec.CommandTimeout = 60
    DBExec.CursorLocation = adUseClient
    DBExec.Open conStr_
    GetDBExecConnec = True
    Exit Function
    
GetDBConnectionErr:
    If i > 3 Then
        GetDBExecConnec = False
        MsgBox "ไม่สามารถติดต่อ Server ได้"
        Exit Function
    End If
    If Err.Number = "-2147467259" Then GoTo RetryCon
    Exit Function
End Function

Private Sub CloseDBExec()
On Error Resume Next
    DBExec.Close
    Set DBExec = Nothing
End Sub
Public Function GetDBConnection(Optional conStr As String) As Boolean
On Error GoTo GetDBConnectionErr:
Dim conStr_ As String, i As Integer
Dim SQLServer
    GetCnString
'    If UCase(ConnectionType) = "INTERNET" Then
'        If testConnection = True Then
'            Exit Function
'        Else
'            Load frmClient
'            frmClient.StateConnect
'            frmClient.CheckState
'            Exit Function
'        End If
'    End If
    If GetDBConnection = True Then Exit Function
    If DBConn.State = adStateOpen Then DBConn.Close
    If conStr <> "" Then
        conStr_ = conStr
    Else
        conStr_ = DBConString
    End If
    
RetryCon:
On Error GoTo GetDBConnectionErr:
    i = i + 1
    GetCnString
'    Dim MsgState As String
'    MsgState = GetSQLState
'    If MsgState <> "" Then
'        MsgBox MsgState
'        GetDBConnection = False
'        Exit Function
'    End If
    Set DBConn = New ADODB.Connection
    DBConn.ConnectionTimeout = 15
    DBConn.CommandTimeout = 60
    DBConn.CursorLocation = adUseClient
    DBConn.ConnectionString = conStr_
    DBConn.Open conStr_
    GetDBConnection = True
    Exit Function
    
GetDBConnectionErr:
    If i > 3 Then
        GetDBConnection = False
        Dim lErrNo As Long
        Dim sErrMSG As String
        Dim sErrSource As String
        With Err
            lErrNo = .Number
            sErrMSG = .Description
            sErrSource = .Source
        End With
        Err.Raise lErrNo, sErrSource, sErrMSG
'        MsgBox "ไม่สามารถติดต่อ Server ได้"
        Exit Function
    End If
    If Err.Number = "-2147467259" Then GoTo RetryCon
    Exit Function
End Function

Public Sub CloseDBConnection()
On Error Resume Next
    DBConn.Close
    Set DBConn = Nothing
End Sub

'Public Function GetSQLState() As String
'On Error GoTo ErrD
'    Dim objWMIService As Object
'    Dim colItems As Object
'    Dim objItem  As Object
'    Dim strMachineName As String
'    Dim StrResult As String
'
'    strMachineName = strSvr
'    StrResult = "ไม่สามารถเชื่อต่อกับเซิฟเวอร์ " & strSvr & " ได้"
'
'    On Error Resume Next
'    If ScanIp(strSvr) = False Then GetSQLState = StrResult: Exit Function
'
'    Exit Function
'ErrD:
'    GetSQLState = Err.Number & vbCrLf & Err.Description
'End Function
'
'Public Function ScanIp(ByVal HostIP As String, Optional ByVal ByteSize As String = "******") As Boolean
'    Dim ECHO As ICMP_ECHO_REPLY
'    Dim pos As Long
'    Dim success As Long
'    Dim txtMsg As String
'    If SocketsInitialize() Then
'
'        success = Ping((HostIP), (ByteSize), ECHO)
'        txtMsg = GetStatusCode(success)
'
'        SocketsCleanup
'
'        If txtMsg = "ip success" Then
'            ScanIp = True
'        Else
'            ScanIp = False
'        End If
'
'    Else
'        ScanIp = False
'        MsgBox "Windows Sockets for 32 bit Windows " & _
'               "environments is not successfully responding."
'
'    End If
'End Function

Public Function GetRS(ByVal tSql As String, Optional ByVal CursorType_ As CursorTypeEnum = 0, Optional ByVal LockType_ As LockTypeEnum = 1, Optional chkConStr As String = "") As ADODB.Recordset
On Error GoTo ErrD
Dim RS As New ADODB.Recordset
Dim strSql As String, i As Integer
Dim f As Form
'    If UCase(ConnectionType) = "INTERNET" Then
'        If testConnection = False Then frmClient.StateConnect: GetRS tSql
'        Call fetchData(tSql, 1)
'        If boolError = False Then
'            Set RS = rsSelect.Clone
''            frmClient.CheckState
'            For Each f In Forms
'                If f.Name = "frmMain" Then
'                    frmMain.Stb1.Panels(3).Text = "Retrieved data " & RS.RecordCount & " record(s)."
'                    Exit For
'                End If
'            Next
'        Else
'            For Each f In Forms
'                If f.Name = "frmMain" Then
'                    frmMain.Stb1.Panels(3).Text = "Error" & vbCrLf & strServerError
'                    Exit For
'                End If
'            Next
'        End If
'        Set GetRS = RS
'        Set RS = Nothing
'        Exit Function
'    Else
'        GoTo Retry
'        Exit Function
'    End If

Retry:
    i = i + 1
    '*** remark here for bwg stock
    If chkConStr = "" Then chkConStr = DBConString
    If GetDBConnection(chkConStr) = False Then Exit Function
    With RS
        .ActiveConnection = DBConn
        .CursorLocation = adUseClient
        .LockType = LockType_
        .CursorType = CursorType_
        .Source = tSql
        .Open
    End With
    Set GetRS = RS
    Set RS.ActiveConnection = Nothing
    Set RS = Nothing
Exit Function

ErrD:
    If i > 3 Then
        If Err.Number = -2147467259 Then
            DoEvents
            MsgBox "ไม่สามารถติดต่อกับเซิฟเวอร์ได้" & vbCrLf & "กรุณาลองเปิดโปรแกรมใหม่"
            Exit Function
        Else
            MsgBox Err.Description & vbCrLf & "กรุณาลองเปิดโปรแกรมใหม่"
            Exit Function
        End If
    Else
        GoTo Retry
    End If
    Set RS = Nothing
End Function

Public Sub AddParameters(ByVal ProcName As String, ByRef colParam() As ParamType)
    Dim cmd As ADODB.Command
    Dim param1() As Parameter
    Dim i As Integer
    Set cmd = New ADODB.Command
    GetDBConnection
    cmd.ActiveConnection = DBConn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "sp_IncomeAnalysisbyCustomer"
    cmd.Parameters.Refresh
    ReDim param1(UBound(colParam))
    For i = 0 To UBound(colParam) - 1
        Dim prm As ADODB.Parameter
        Set prm = cmd.CreateParameter(colParam(i).ParamName, adVarChar, adParamInput, 100, colParam(i).ParamValue)
        cmd.Parameters.Append prm
    Next i
    cmd.Execute
    Set cmd = Nothing
End Sub
