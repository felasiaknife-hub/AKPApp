VERSION 5.00
Begin VB.Form frmFile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File Picker"
   ClientHeight    =   2160
   ClientLeft      =   2220
   ClientTop       =   3540
   ClientWidth     =   4185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   2304
      TabIndex        =   2
      Top             =   324
      Width           =   1848
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   324
      TabIndex        =   1
      Top             =   648
      Width           =   1920
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   72
      TabIndex        =   0
      Top             =   324
      Width           =   2208
   End
   Begin VB.Label lblCol 
      Caption         =   "Label4"
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   336
   End
   Begin VB.Label lblRow 
      Caption         =   "Label4"
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   336
   End
   Begin VB.Label lblFile 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   192
      Left            =   108
      TabIndex        =   3
      Top             =   72
      Width           =   4032
   End
End
Attribute VB_Name = "frmFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Dir1_Change()
    File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1
End Sub


Private Sub File1_Click()
    lblFile = File1.Path & "\" & File1.filename
    Tag = File1.filename
End Sub


Private Sub File1_DblClick()
    Hide
End Sub


Private Sub Form_Deactivate()
    
    ' update grid value
    If Tag <> "" Then frmMain.fa.Cell(flexcpText, lblRow, lblCol) = Tag
    
    ' go away
    Hide

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    ' esc clears, then quits
    If KeyAscii = 27 Then lblFile = "": Hide
    
    ' enter quits
    If KeyAscii = 13 Then Hide

End Sub

