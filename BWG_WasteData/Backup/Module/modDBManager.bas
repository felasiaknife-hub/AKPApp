Attribute VB_Name = "modDBManager"
Option Explicit
'Public Const DBConString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Krisana\Projects\AWS_Stock\aws_stock.mdb;Persist Security Info=False"
Public DBConString As String
Public DBManager As AWSDBManager
Public DBConn As New ADODB.Connection
Public MoveWaste As String
Public Years_  As String
Public WasteID As String
Public Wastelab_ID As String
Public WasteReferent As String
Public countID As String
'Public Const DSN_NAME As String = "DSN_MMS"
'Public Function GetDBConnection(Optional conStr As String) As Boolean
'On Error GoTo GetDBConnectionErr:
'Dim conStr_ As String, i As Integer
'    GetDBConnection = True
'    If DBConn.State = adStateOpen Then DBConn.Close
'    If conStr <> "" Then
'        conStr_ = conStr
'    Else
'        conStr_ = DBConString
'    End If
'
'RetryCon:
'    i = i + 1
'    Set DBConn = New ADODB.Connection
'    DBConn.ConnectionTimeout = 60
'    DBConn.CommandTimeout = 180
'    DBConn.CursorLocation = adUseServer
'    DBConn.Open conStr_
'    Exit Function
'
'GetDBConnectionErr:
'    If i > 3 Then
'        GetDBConnection = False
'        MsgBox "ไม่สามารถติดต่อ Server ได้"
'        Exit Function
'    End If
'    If Err.Number = "-2147467259" Then GoTo RetryCon
'    Exit Function
'End Function
