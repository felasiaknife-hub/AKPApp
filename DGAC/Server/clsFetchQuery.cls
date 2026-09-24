VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFetchQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim intTraffic As Long
Public intRecs As Long
Public intX As Integer, intY As Integer
Public strFieldNames As String
Public strRecs As String
Public rsCls As Recordset

Public Sub getSortDataNSend(ByVal Index As Integer, ByVal strData As String)
On Error GoTo err
intTraffic = 0
Set rsCls = New ADODB.Recordset
boolCancelled = False

If Left(strData, 1) = "1" Then
    rsCls.Open Mid(strData, 3, Len(strData)), dbCnn1, adOpenDynamic, adLockOptimistic
ElseIf Left(strData, 1) = "2" Then
    rsCls.Open Mid(strData, 3, Len(strData)), dbCnn2, adOpenDynamic, adLockOptimistic
ElseIf Left(strData, 1) = "3" Then
    rsCls.Open Mid(strData, 3, Len(strData)), dbCnn2, adOpenDynamic, adLockOptimistic
Else
    Exit Sub
End If


strFieldNames = ""
strRecs = ""
intRecs = 0
intX = 0
intY = 0
While intX < rsCls.Fields.Count
    strFieldNames = strFieldNames & "#^" & rsCls.Fields(intX).Name & "," & _
    rsCls.Fields(intX).DefinedSize + 1 '+1 To make sure no data is lost and to avoid
                                    'multi step error in client
    intX = intX + 1
Wend
'sending the header fields to client to prepare the recordset
frmServer.sckServer(Index).sendData "#^SELECTHEADER^#" & strFieldNames & "#^EOF^#"
Call Wait(0.1) 'wait till the header is sent accross

If rsCls.EOF = False Then
    rsCls.MoveFirst
    While rsCls.EOF = False
        intX = 0
        While intX < rsCls.Fields.Count
            strRecs = strRecs & "#^" & rsCls(intX)
            intX = intX + 1
        Wend
        rsCls.MoveNext
        intRecs = intRecs + 1
        If intRecs Mod 1000 = 0 Then
            intTraffic = intTraffic + 1
            Call Wait(0.1)
            If boolCancelled = True Then
                Call AddSeverMsg("Query cancelled.")
                Exit Sub
            ElseIf frmServer.sckServer(Index).State = sckClosed Or frmServer.sckServer(Index).State = sckClosing Then
                Exit Sub
            Else
                Call AddSeverMsg("Experiencing class : " & intTraffic & " traffic from " & frmServer.sckServer(Index).RemoteHostIP)
            End If
        End If
    Wend
    'send the data stored in the string to client
    frmServer.sckServer(Index).sendData "#^SELECT^#" & strRecs
    frmServer.sckServer(Index).sendData "#^EOF^#"
    Call AddSeverMsg(intRecs & " record(s) retrieved.")
Else
    Call AddSeverMsg("No records retrieved.")
    frmServer.sckServer(Index).sendData "#^SELECT^##^EOF^#"
End If
Exit Sub

err:
If err.Number = 340 Then
    Call AddSeverMsg("Client disconnected")
    Exit Sub
Else
    frmServer.sckServer(Index).sendData "#^ERROR^#" & err.Description
    Call AddSeverMsg(err.Description)
End If
End Sub


