VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1020
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
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
      Height          =   675
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StrTitle As String
Public Msg As String
Public TimeEnable As Boolean
Public TimeSecond As Integer
Public cmd1Visible As Boolean, cmd2Visible As Boolean, cmd3Visible As Boolean
Public cmd1Text As String, cmd2Text As String, cmd3Text As String
Public tSql As String

Public frmOwner As Form
Public CancelState As String

Dim SS As Integer
Dim MS As Long

Private Sub Command3_Click()
    If UCase(CancelState) = "END" Then Unload frmClient: End
    If UCase(CancelState) = "HIDE" Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Public Sub SetParameter()
    ProgressBar1.Value = 0
    ProgressBar1.Min = 0
    ProgressBar1.Max = 1
    Me.Caption = StrTitle
    Label1.Caption = Msg
    
    Timer1.Enabled = TimeEnable
    Timer2.Enabled = TimeEnable
    If TimeEnable = False Then
        Unload Me
        Exit Sub
    End If
    
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
    If Label1.Caption = Msg Then
        Label1.Caption = ""
    Else
        Label1.Caption = Msg
    End If
    
    ProgressBar1.Max = MS + 1
    ProgressBar1.Value = MS
    
    If testConnection = True Then
        Unload Me
    Else
        frmClient.StateConnect
        Unload Me
    End If
End Sub

