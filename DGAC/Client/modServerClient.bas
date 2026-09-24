Attribute VB_Name = "modServerClient"
Option Explicit

Public strServerError As String
Public boolFileExistOnServer As Boolean
Public arrFile() As String
Public FSO
'Timer
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
'ends

'kill X button
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Const MF_BYPOS = &H400&
'ends

'timeout
Public Const lngTimeOut = 10 'Minutes
Public lngIdleTime As Long
'ends

Public boolQueryExecuted As Boolean
Public boolQueryExecuting As Boolean
Public strFields() As String
Public arrData() As String
Public rsSelect As ADODB.Recordset
Public boolError As Boolean
Public strNowDateTime As String
Public strData As String
Public strHeader As String

Public Sub KillCloseButton(hwnd As Long)
Dim hSysMenu As Long
hSysMenu = GetSystemMenu(hwnd, 0)
Call RemoveMenu(hSysMenu, 6, MF_BYPOS)
Call RemoveMenu(hSysMenu, 5, MF_BYPOS)
End Sub

'This gives enough time to fill in the records / data
Public Function Wait(ByVal TimeToWait As Double) 'Time In seconds
Dim EndTime As Long
EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
Do Until GetTickCount > EndTime
    DoEvents
Loop
End Function

Public Function chkFilePath(strFilePath As String) As Boolean
If (FSO.fileexists(strFilePath)) Then
    chkFilePath = True
Else
    chkFilePath = False
End If
End Function

Public Function GetFileName(strFilePath As String) As String
If chkFilePath(strFilePath) = True Then
    arrFile = Split(strFilePath, "\")
    GetFileName = arrFile(UBound(arrFile))
Else
    GetFileName = ""
End If
Exit Function

err:
GetFileName = ""
End Function

Public Function chkFolderPath(strFilePath As String, boolCreateFolder As Boolean) As Boolean
If (FSO.folderexists(strFilePath)) Then
    chkFolderPath = True
Else
    chkFolderPath = False
    If boolCreateFolder = True Then
        On Error GoTo err:
        FSO.CreateFolder (strFilePath)
        chkFolderPath = True
    End If
End If
Exit Function

err:
chkFolderPath = False
End Function

Public Function GetVirtualFileName(strFilePath As String) As String
If Len(strFilePath) > 0 Then
    arrFile = Split(strFilePath, "\")
    GetVirtualFileName = arrFile(UBound(arrFile))
Else
    GetVirtualFileName = ""
End If
Exit Function

err:
GetVirtualFileName = ""
End Function

'testing the existence of connection
Public Function testConnection() As Boolean
On Error GoTo err:
If frmClient.Winsock1.State = sckClosed Or _
    frmClient.Winsock1.State = sckClosing Or _
    frmClient.Winsock1.State = sckError Or _
    frmClient.Winsock1.State = sckConnecting Then
        
    GoTo err
Else
    testConnection = True
End If
Exit Function

err:
    testConnection = False
    boolError = True
    boolQueryExecuting = False
    MsgBox "Error in connecting to server.", vbCritical
End Function

'checking date and time
Public Function checkDateTime() As String
If testConnection = True Then
    frmClient.Winsock1.SendData "#^DATETIME^#"
    strNowDateTime = ""
    Call Wait(0.1)
End If
End Function

'check data, parse and fill it
Sub chkData()
'On Error GoTo err
boolError = False
If InStr(1, strData, "#^SELECT^#") <> 0 Then
    strHeader = "#^SELECT^#"
ElseIf InStr(1, strData, "#^SELECTHEADER^#") <> 0 Then
    strHeader = "#^SELECTHEADER^#"
ElseIf InStr(1, strData, "#^EXECUTE^#") <> 0 Then
    strHeader = "#^EXECUTE^#"
ElseIf InStr(1, strData, "#^FILE^#") <> 0 Then
    strHeader = "#^FILE^#"
ElseIf InStr(1, strData, "#^DATETIME^#") <> 0 Then
    strHeader = "#^DATETIME^#"
ElseIf InStr(1, strData, "#^ERROR^#") <> 0 Then
    strHeader = "#^ERROR^#"
End If

Dim intY As Long
Dim intX As Long
Dim strA() As String

' set parser
If strHeader = "#^SELECTHEADER^#" Then
    Set rsSelect = Nothing
    Set rsSelect = New ADODB.Recordset
    rsSelect.CursorLocation = adUseClient
    rsSelect.LockType = adLockOptimistic
    intY = 1
    
    strA = Split(Replace(strData, "#^SELECTHEADER^#", ""), "#^")
    While intY <= UBound(strA)
        strFields = Split(strA(intY), ",")
        If strFields(1) > 20 Then
            rsSelect.Fields.Append strFields(0), adVarChar, strFields(1)
        Else
            rsSelect.Fields.Append strFields(0), adVarChar, 50
        End If
        intY = intY + 1
    Wend
    rsSelect.Open
    frmAccessing.Show
    strData = ""
ElseIf strHeader = "#^SELECT^#" Then
    intY = 0
    intX = 1
    strA = Split(Replace(strData, "#^SELECT^#", ""), "#^")
    While intX <= UBound(strA)
        rsSelect.AddNew
        While intY < rsSelect.Fields.Count
            rsSelect.Fields(intY) = strA(intX)
            intY = intY + 1
            intX = intX + 1
        Wend
        rsSelect.Update
        intY = 0
    Wend
    strData = ""
    Unload frmAccessing
ElseIf strHeader = "#^EXECUTE^#" Then
    'already handled on data arrival so blank
ElseIf strHeader = "#^FILE^#" Then
    
ElseIf strHeader = "#^DATETIME^#" Then
    strNowDateTime = Replace(strData, "#^DATETIME^#", "")
    If Len(strNowDateTime) <= 0 Then
        Call checkDateTime
    End If
    strData = ""
ElseIf strHeader = "#^ERROR^#" Then
    boolError = True
End If
Exit Sub

err:
    strData = ""
    Unload frmAccessing
    boolError = True
    Exit Sub
End Sub

Public Sub fetchData(strQuery As String, DBId As Integer)
If testConnection = True Then
    If UCase(Left(strQuery, 6)) = "SELECT" Then
        On Error GoTo err:
        frmClient.Winsock1.SendData "#^SELECT^#" & DBId & "#" & strQuery
        boolQueryExecuting = True
        While boolQueryExecuting = True
            Call Wait(0)
        Wend
    Else
        MsgBox "The query type should be SELECT for this command.", vbExclamation
        boolError = True
        Exit Sub
    End If
End If
Exit Sub

err:
MsgBox err.Description, vbExclamation
boolQueryExecuting = False
End Sub

Public Sub executeQuery(strQuery As String, DBId As Integer)
On Error GoTo err
boolQueryExecuted = False
If testConnection = True Then
    If UCase(Left(strQuery, 6)) = "INSERT" Or UCase(Left(strQuery, 6)) = "UPDATE" Or UCase(Left(strQuery, 6)) = "DELETE" Then
        frmClient.Winsock1.SendData "#^EXECUTE^#" & DBId & "#" & strQuery
        boolQueryExecuting = True
        While boolQueryExecuting = True
            Call Wait(0.2)
        Wend
    Else
        MsgBox "The query type should be INSERT, UPDATE or DELETE for this command.", vbExclamation
        boolError = True
        Exit Sub
    End If
End If
Exit Sub

err:
MsgBox err.Description, vbExclamation
boolQueryExecuting = False
End Sub

Public Function fileExistsOnServer(FilePath As String) As Boolean
On Error GoTo err
boolFileExistOnServer = False
If testConnection = True Then
    frmClient.Winsock1.SendData "#^FILEEXISTS^#" & FilePath
    Call Wait(0.1)
    If boolError = False Then
        fileExistsOnServer = boolFileExistOnServer
    End If
End If
Exit Function

err:
MsgBox err.Description, vbExclamation
boolQueryExecuting = False
End Function
