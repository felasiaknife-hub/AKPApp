VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3585
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   3585
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   600
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer TimerState 
      Interval        =   10000
      Left            =   1800
      Top             =   720
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   2580
      TabIndex        =   4
      Top             =   540
      Width           =   915
   End
   Begin VB.TextBox txtPort 
      Height          =   345
      Left            =   2580
      TabIndex        =   2
      Text            =   "1001"
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox txtServer 
      Height          =   345
      Left            =   780
      TabIndex        =   1
      Text            =   "192.168.1.7"
      Top             =   120
      Width           =   1155
   End
   Begin VB.Timer timerTimeOut 
      Interval        =   60000
      Left            =   1260
      Top             =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Port"
      Height          =   195
      Left            =   2220
      TabIndex        =   3
      Top             =   180
      Width           =   315
   End
   Begin VB.Label Label3 
      Caption         =   "Server"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   555
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub StateConnect()
    If Winsock1.State = 0 Or Winsock1.State = 9 Then
        Winsock1.Connect strSvr, ConnectionPort
        Set FSO = CreateObject("Scripting.FileSystemObject")
        CheckState
    End If
End Sub

Private Sub cmdClose_Click()
    Form_Load
    Me.Hide
End Sub

Private Sub Form_Initialize()
    txtServer.Text = strSvr
    txtPort.Text = ConnectionPort
End Sub

Private Sub Form_Load()
    'xx
End Sub

Private Sub TimerState_Timer()
    CheckState
End Sub

Private Sub timerTimeOut_Timer()
lngIdleTime = lngIdleTime + 1
If lngIdleTime >= lngTimeOut Then
    Winsock1.Close
    MsgBox "Your session has expired. Please login to continue", vbExclamation
End If
End Sub

Public Sub CheckState()
    Dim FormMainLoaded As Boolean
    Dim f As Form
    For Each f In Forms
        If f.Name = "frmMain" Then
            FormMainLoaded = True
            frmMain.Stb1.Panels(3).Text = WinsockState(Winsock1.State)
            Exit Sub
        End If
    Next
    
    If testConnection = True Then Exit Sub
    If FormMainLoaded = False Then
        frmMsg.Hide
        Load frmMsg
        With frmMsg
            .StrTitle = "Check server status"
            .msg = WinsockState(Winsock1.State)
            .TimeEnable = True
            .cmd1Text = ""
            .cmd2Text = ""
            .cmd3Text = "Cancel"
            .cmd1Visible = False
            .cmd2Visible = False
            .cmd3Visible = True
            .CancelState = "END"
            .SetParameter
            .Show 1
        End With
    End If
End Sub

Private Function WinsockState(ByVal State As Integer) As String
    Select Case State
        Case 0: WinsockState = "Connection closed"
        Case 1: WinsockState = "Open"
        Case 2: WinsockState = "Listening for incoming connections"
        Case 3: WinsockState = "Connection pending"
        Case 4: WinsockState = "Resolving remote host name"
        Case 5: WinsockState = "Remote host name successfully resolved"
        Case 6: WinsockState = "Connecting to remote host"
        Case 7: WinsockState = "Connected to remote host"
        Case 8: WinsockState = "Connection Is closing"
        Case 9: WinsockState = "Error occured"
    End Select
End Function

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
