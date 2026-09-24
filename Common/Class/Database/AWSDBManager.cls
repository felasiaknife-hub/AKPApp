VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AWSDBManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim conStr_ As String
Dim curConn_ As ADODB.Connection

Public Function GetDBConnection(Optional conStr As String) As ADODB.Connection
On Error GoTo ConnectErr:
    If Not curConn_ Is Nothing Then
        If (LCase(Trim(conStr_)) = LCase(Trim(conStr)) Or Trim(conStr) = vbNullString) And curConn_.State = adStateOpen Then
            Set GetDBConnection = curConn_
            Exit Function
        ElseIf curConn_.State = adStateOpen And Trim(conStr) <> vbNullString Then
            CloseConnection
        End If
    End If

    If curConn_.State = adStateOpen Then CloseConnection
    If Trim(conStr) <> vbNullString Then
        conStr_ = conStr
    End If
    
    Set curConn_ = New ADODB.Connection
    If curConn_.State = adStateOpen Then curConn_.Close
    curConn_.CommandTimeout = 60 * 10
    curConn_.ConnectionTimeout = 60
    curConn_.CursorLocation = adUseClient
    curConn_.open conStr_
    
    Set GetDBConnection = curConn_
    Exit Function

RetryConn:
On Error Resume Next
    Set curConn_ = New ADODB.Connection
    If curConn_.State = adStateOpen Then curConn_.Close
    curConn_.CommandTimeout = 60 * 10
    curConn_.ConnectionTimeout = 60
    curConn_.CursorLocation = adUseClient
    curConn_.open conStr_
    
    Set GetDBConnection = curConn_
    Exit Function
    
ConnectErr:
    CloseConnection
    Set GetDBConnection = Nothing
'    MsgBox Err.Description, vbCritical, "Error"
    GoTo RetryConn
    Exit Function
End Function

Public Function CloseConnection() As Boolean
On Error GoTo CloseConnectionErr:
    If Not curConn_ Is Nothing Then
        If curConn_.State = adStateOpen Then
            curConn_.Close
            Set curConn_ = Nothing
        End If
    End If
    CloseConnection = True
    Exit Function

CloseConnectionErr:
    CloseConnection = False
    Set curConn_ = Nothing
    Exit Function
End Function

Public Property Get ConnectionString() As String
    ConnectionString = conStr_
End Property

Public Property Let ConnectionString(ByVal newVal As String)
    conStr_ = newVal
End Property
