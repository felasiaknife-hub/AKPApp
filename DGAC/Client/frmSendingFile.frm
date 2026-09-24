VERSION 5.00
Begin VB.Form frmSendingFile 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3900
   Icon            =   "frmSendingFile.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copying File ... Please Wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2550
   End
End
Attribute VB_Name = "frmSendingFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim f As Form

Private Sub Form_Load()
KillCloseButton Me.hwnd

For Each f In Forms
    If f.Name <> Me.Name Then
        f.Enabled = False
    End If
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
For Each f In Forms
    If f.Name <> Me.Name Then
        f.Enabled = True
    End If
Next
End Sub

Private Sub Timer1_Timer()
If Label1.Caption = "Copying File ... Please Wait..." Then
    Label1.Caption = ""
Else
    Label1.Caption = "Copying File ... Please Wait..."
End If
If testConnection = False Then
    Unload Me
End If
End Sub
