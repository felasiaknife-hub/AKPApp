VERSION 5.00
Begin VB.Form frmFlexNote 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Flex Cell Note"
   ClientHeight    =   2850
   ClientLeft      =   2940
   ClientTop       =   3390
   ClientWidth     =   2595
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   2595
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   855
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2115
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmFlexNote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' frmFlexNote
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Use this form to add Excel-style cell notes to VSFlexGrid controls.
'
' The parent grid must handle the MouseMove event to create an instance of
' this form and call the form's ShowNote method.
'
' Note the we use variables of type "Control" throughout. This allows
' the same form to be used with any flavor of VSFlexGrid (ADO, DAO, etc)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private m_Ctl As Control
Private m_lRow As Long
Private m_lCol As Long

Private Sub Form_Deactivate()

    ' if we have a note, save it
    If Len(txtNote) Then
        m_Ctl.Cell(flexcpData, m_lRow, m_lCol) = txtNote.Text
    
    ' no text? no note!
    Else
        Set m_Ctl.Cell(flexcpData, m_lRow, m_lCol) = Nothing
        Set m_Ctl.Cell(flexcpPicture, m_lRow, m_lCol) = Nothing
    End If
    Unload Me
End Sub

Private Sub Form_Resize()
    txtNote.Move 0, 0, ScaleWidth - 40, ScaleHeight - 40
End Sub

Public Sub ShowNote(ctl As Control, row As Long, col As Long)

    ' save control info to update note later
    Set m_Ctl = ctl
    m_lRow = row
    m_lCol = col

    ' copy font from parent control to look nice
    Set txtNote.Font = ctl.Font
    Set lblNote.Font = ctl.Font
    
    ' calculate note position
    Dim fLeft!, fTop!, fWid!, fHei!
    With ctl
        fLeft = .Parent.Left + .Left + .ColPos(col) + .ColWidth(col) + 200
        fTop = .Parent.Top + .Top + .RowPos(row) + 300
    End With
    
    ' calculate note size
    lblNote = txtNote
    fWid = lblNote.Width + 300
    fHei = lblNote.Height + 150
    
    ' show note (we stay up until deactivated)
    Move fLeft, fTop, fWid, fHei
    txtNote.SelStart = 32000
    Visible = True
End Sub

Private Sub txtNote_Change()

    ' resize note as the user types
    Dim fWid!, fHei!
    lblNote = txtNote
    fWid = lblNote.Width + 300
    fHei = lblNote.Height + 150
    Move Left, Top, fWid, fHei
End Sub

Private Sub txtNote_KeyPress(KeyAscii As Integer)

    ' quit without saving when user hits Escape
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub
