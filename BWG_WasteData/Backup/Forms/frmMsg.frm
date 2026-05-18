VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message Box"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   7665
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   900
      Top             =   1620
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   300
      Top             =   1620
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   1620
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   1620
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Label Label1 
      Height          =   1395
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Msg As String
Public TimeEnable As Boolean
Public TimeSecond As Integer
Public cmd1Visible As Boolean, cmd2Visible As Boolean, cmd3Visible As Boolean
Public cmd1Text As String, cmd2Text As String, cmd3Text As String

Public frmOwner As Form

Dim SS As Integer

Dim MS As Long

Private Sub Form_Load()
    Timer2.Enabled = True
End Sub


Private Sub Timer1_Timer()
    SS = SS + 1
    DBManager.GetDBConnection.Close
    Debug.Print DBManager.GetDBConnection.State
    If DBManager.GetDBConnection.State = 1 Then
        Label1.Caption = "ติดต่อกับ Server ได้แล้ว"
        ConnectState = True
        Unload Me
    End If
'    If SS = TimeSecond Then
'        Unload Me
'    End If
End Sub

Public Sub SetParameter()
    Timer1.Enabled = TimeEnable
    Label1.Caption = Msg
    Command1.Visible = cmd1Visible
    Command2.Visible = cmd2Visible
    Command3.Visible = cmd3Visible
    Command1.Caption = cmd1Text
    Command2.Caption = cmd2Text
    Command3.Caption = cmd3Text
End Sub

Private Sub Timer2_Timer()
Dim i As Long
Dim StrX As String
    MS = MS + 1
    For i = 1 To MS
        StrX = StrX & ">"
    Next
    DoEvents
    Label1.Caption = Msg & vbCrLf & StrX & vbCrLf & vbCrLf & "** ถ้าหากท่านรอนานเกิน 10 นาที กรุณาปิดโปรแกรมแล้วเปิดใหม่ หรือติดต่อเจ้าหน้าที่ IT..."
End Sub


