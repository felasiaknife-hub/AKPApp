VERSION 5.00
Begin VB.Form frmAccessing 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   3255
   Icon            =   "frmAccessing.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1030
      TabIndex        =   1
      Top             =   790
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Accessing Data ... Please Wait..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2850
   End
End
Attribute VB_Name = "frmAccessing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim f As Form

Private Sub cmdCancel_Click()
    frmClient.Winsock1.SendData "#^CANCELSELECT^#"
    boolError = True
    Unload Me
End Sub

Private Sub Form_Load()
    boolQueryExecuting = True
    
    KillCloseButton Me.hwnd
    
    For Each f In Forms
        If f.Name <> Me.Name Then
            f.Enabled = False
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    boolQueryExecuting = False
    For Each f In Forms
        If f.Name <> Me.Name Then
            f.Enabled = True
        End If
    Next
End Sub

Private Sub Timer1_Timer()
    If Label1.Caption = "Accessing Data ... Please Wait..." Then
        Label1.Caption = ""
    Else
        Label1.Caption = "Accessing Data ... Please Wait..."
    End If
    If testConnection = False Then
        Unload Me
    End If
End Sub
