VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Form1 
   Caption         =   "VSFlexGrid Demo: Cell Notes"
   ClientHeight    =   4755
   ClientLeft      =   2010
   ClientTop       =   1875
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   ScaleHeight     =   4755
   ScaleWidth      =   8385
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   3855
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   8115
      _cx             =   14314
      _cy             =   6800
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FormatCells     =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label lbl 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   9135
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   0
      Picture         =   "Form1.frx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' fg_MouseMove
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
' This routine show how to use the frmFlexNote form type to add
' Excel-style cell notes to VSFlexGrid controls.
'
' To use this routine in other projects, remember to add the the
' frmFlexNote.frm file to the project.
'
' Like Excel, this form uses little red markers to indicate cells with
' notes attached to them. While the note is visible, the user may edit
' it. The note will resize as the user types to accommodate the text,
' and the changes will be saved back into the cell's flexcpData property.
' To cancel edits made to a note, hit escape. To accept the edits, just
' click another cell or another control.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private Sub fg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static lNoteRow&, lNoteCol&, r&, c&

    ' clicking? no work
    If Button <> 0 Then Exit Sub
    
    ' get mouse coordinates
    r = fg.MouseRow
    c = fg.MouseCol
    
    ' same cell or neighbour? no work
    If r = lNoteRow And c = lNoteCol Then Exit Sub
    If r = lNoteRow And c = lNoteCol + 1 Then Exit Sub
    
    ' other cell, hide current note, if any
    If lNoteRow >= 0 And lNoteCol >= 0 Then
        fg.SetFocus
        lNoteRow = -1
        lNoteCol = -1
    End If
    
    ' no note to show? then bail out
    If r <= 0 Or c <= 0 Then Exit Sub
    If TypeName(fg.Cell(flexcpData, r, c)) <> "String" Then Exit Sub
    
    ' show new note
    Dim nt As New frmFlexNote
    nt.txtNote = fg.Cell(flexcpData, r, c)
    nt.ShowNote fg, r, c
    
    ' save coordinates for next time
    lNoteRow = r
    lNoteCol = c

End Sub

Private Sub fg_DblClick()

    ' create a new note here
    Dim r&, c&
    r = fg.row
    c = fg.col
    If TypeName(fg.Cell(flexcpData, r, c)) = "String" Then Exit Sub
    fg.Cell(flexcpData, r, c) = "** New Note **" & vbCrLf
    fg.Cell(flexcpPicture, r, c) = imgNote
    fg.Cell(flexcpPictureAlignment, r, c) = flexPicAlignRightTop
    
End Sub

Private Sub Form_Load()

    ' show a little help
    lbl = "Cells with a red mark have notes attached to them. Move the mouse over the cell to see the note." & vbCrLf & _
          "You may edit the notes by typing into the note or create new notes by double-clicking any cell."

    ' make the grid editable
    fg.Editable = True
    fg.HighLight = flexHighlightWithFocus
    
    ' create some notes at random
    ' notes are simply strings attached to a cell's 'flexcpData' property.
    Dim i&, r&, c&
    For i = 1 To 20
        r = Rnd * (fg.Rows - 2) + 1
        c = Rnd * (fg.Cols - 2) + 1
        fg.Cell(flexcpData, r, c) = "** Note " & i & " **" & vbCrLf & "I'm a note, just a little note." & vbCrLf & "Edit me if you want."
        fg.Cell(flexcpPicture, r, c) = imgNote
        fg.Cell(flexcpPictureAlignment, r, c) = flexPicAlignRightTop
    Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fg.Move fg.Left, fg.Top, ScaleWidth - 2 * fg.Left, ScaleHeight - fg.Top - fg.Left
    lbl.Width = fg.Width
End Sub
