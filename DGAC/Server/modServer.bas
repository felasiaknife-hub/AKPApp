Attribute VB_Name = "modServer"
Option Explicit

Public arrFile() As String
Public FSO

'wait
    Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
'ends

'------------------Change Date Format------------------
Public Const LOCALE_SSHORTDATE = &H1F
Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
Public LCID As Long
'------------------ENDS------------------

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Public dbCnn1 As ADODB.Connection
Public dbCnn2 As ADODB.Connection
Public dbCnn3 As ADODB.Connection

Public boolPassword As Boolean
Public boolCancelled As Boolean
Dim intList As Integer
Public rs As ADODB.Recordset

'This gives enuf time to fill in the records / data
Public Function Wait(ByVal TimeToWait As Double) 'Time In seconds
Dim EndTime As Double
EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
Do Until GetTickCount > EndTime
    DoEvents
Loop
End Function

Sub changeDateFormat()
If CStr(Format(Now(), "Short Date")) <> CStr(Format(Now(), "dd-MMM-yyyy")) Then
    LCID = GetSystemDefaultLCID()
    Call SetLocaleInfo(LCID, LOCALE_SSHORTDATE, "dd-MMM-yyyy")
    Call AddSeverMsg("Set date format.")
End If
End Sub

Function validateNumber(txtChk As TextBox, KeyAscii As Integer) As Boolean
validateNumber = False
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
    KeyAscii = 0
End If
validateNumber = True
End Function

'make sure the data is coming from the allowed clients only
Function chkClient(Index As Integer) As Boolean
chkClient = False
intList = 0
While intList <= frmServer.lstClients.ListCount
    If frmServer.sckServer(Index).RemoteHostIP = frmServer.lstClients.List(intList) Then
        chkClient = True
        Exit Function
    End If
    intList = intList + 1
Wend
End Function

'to auto add the string to the server msg txt box
Sub AddSeverMsg(strMsg As String)
    frmServer.txtMsg.Text = frmServer.txtMsg.Text & vbCrLf & strMsg
    frmServer.txtMsg.SelStart = Len(frmServer.txtMsg.Text)
End Sub

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

Public Function chkFilePath(strFilePath As String) As Boolean
If (FSO.fileexists(strFilePath)) Then
    chkFilePath = True
Else
    chkFilePath = False
End If
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

