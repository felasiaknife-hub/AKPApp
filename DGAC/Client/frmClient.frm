VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   1995
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   1995
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer timerTimeOut 
      Interval        =   60000
      Left            =   1080
      Top             =   120
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   480
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Winsock1.Connect "58.10.189.98", "55903"
Set FSO = CreateObject("Scripting.FileSystemObject")
End Sub

Private Sub timerTimeOut_Timer()
lngIdleTime = lngIdleTime + 1
If lngIdleTime >= lngTimeOut Then
    Winsock1.Close
    MsgBox "Your session has expired. Please login to continue", vbExclamation
End If
End Sub

Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
lngIdleTime = 0
Dim strTempData As String
Winsock1.GetData strTempData, vbString
strServerError = ""
boolError = False
If InStr(1, strTempData, "#^ERROR^#") <> 0 Then
    boolError = True
    boolQueryExecuting = False
    strServerError = Replace(strTempData, "#^ERROR^#", "")
    strData = ""
    Exit Sub
ElseIf InStr(1, strTempData, "#^PASSWORD^#") <> 0 Then
    Winsock1.SendData "#^PASSWORD^#" & Encrypt("HawkeR", strPassword)
    boolQueryExecuted = True
    boolQueryExecuting = False
    Exit Sub
ElseIf InStr(1, strTempData, "#^EXECUTE^#") <> 0 Then
    boolQueryExecuted = True
    boolQueryExecuting = False
    Exit Sub
ElseIf InStr(1, strTempData, "#^FILEDELETED^#") <> 0 Then
    boolQueryExecuting = False
    boolFileDeleted = True
    Exit Sub
ElseIf InStr(1, strTempData, "#^FILENOTDELETED^#") <> 0 Then
    boolQueryExecuting = False
    boolFileDeleted = False
    Exit Sub
ElseIf InStr(1, strTempData, "#^FILEEXISTS^#") <> 0 Then
    boolFileExistOnServer = True
    boolQueryExecuting = False
    Exit Sub
ElseIf InStr(1, strTempData, "#^FILENOTEXISTS^#") <> 0 Then
    boolFileExistOnServer = False
    boolQueryExecuting = False
    Exit Sub
ElseIf InStr(1, strTempData, "#^FILEBEGIN^#") <> 0 Then
    boolFileExistOnServer = False
    boolQueryExecuting = False
    boolStartDownload = True
    Exit Sub
ElseIf InStr(1, strTempData, "#^FILE^#") <> 0 Then
    'Put data to file
    strData = Replace(strTempData, "#^FILE^#", "")
    Put lngFreeFileNo, , strData
    boolTransfering = True
ElseIf InStr(1, strTempData, "#^FILEEND^#") <> 0 Then
    'End of file sent by server
    Unload frmAccessing
    boolTransfering = False
    Close lngFreeFileNo
ElseIf InStr(1, strTempData, "#^EOF^#") = 0 Then
    strData = strData & strTempData
    Call Wait(0)
ElseIf InStr(1, strTempData, "#^EOF^#") = 0 Then
    strData = strData & strTempData
    Call Wait(0)
ElseIf InStr(1, strTempData, "#^EOF^#") <> 0 Then
    strData = strData & strTempData
    strData = Replace(strData, "#^EOF^#", "")
    Call chkData
    Call Wait(0)
End If
End Sub

Private Sub Winsock1_SendComplete()
lngIdleTime = 0
End Sub
