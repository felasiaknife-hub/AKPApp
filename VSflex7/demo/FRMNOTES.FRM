VERSION 5.00
Begin VB.Form frmNotes 
   BorderStyle     =   0  'None
   Caption         =   "Note Editor"
   ClientHeight    =   1845
   ClientLeft      =   4020
   ClientTop       =   4305
   ClientWidth     =   3705
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNote 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label lblRow 
      Caption         =   "Label4"
      Height          =   300
      Left            =   3360
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lblCol 
      Caption         =   "Label4"
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Deactivate()
    
    ' update grid value
    If Tag <> "" Then frmNew6!fa.Cell(flexcpText, lblRow, lblCol) = Tag
    
    ' go away
    Hide

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    ' esc clears, then quits
    If KeyAscii = 27 Then Tag = "": Hide
    
    ' enter quits
    If KeyAscii = 13 Then Hide


End Sub


Private Sub Form_Resize()

    txtNote.Move 0, 0, ScaleWidth, ScaleHeight

End Sub


Private Sub txtNote_Change()

    Tag = txtNote
    
End Sub


