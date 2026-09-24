VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Form1 
   Caption         =   "FlexGrid Demo: Selection and Scrolling control"
   ClientHeight    =   4605
   ClientLeft      =   1800
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   6585
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   3435
      Left            =   60
      TabIndex        =   6
      Top             =   780
      Width           =   6195
      _cx             =   10927
      _cy             =   6059
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
   Begin VB.CheckBox chkCurrent 
      Caption         =   "Keep &Current Cell Visible"
      Height          =   255
      Left            =   3540
      TabIndex        =   5
      Top             =   480
      Width           =   2895
   End
   Begin VB.CheckBox chkEditing 
      Caption         =   "While &Editing"
      Height          =   255
      Left            =   1620
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.CheckBox chkVertical 
      Caption         =   "&Vertical Ranges Only"
      Height          =   255
      Left            =   3540
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.CheckBox chkProtect 
      Caption         =   "Protect Green &Area"
      Height          =   255
      Left            =   1620
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Restrict Scrolling:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1230
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Restrict Selection:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   1290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkProtect_Click()
    fg.Select fg.Row, fg.Col
End Sub

Private Sub chkVertical_Click()
    fg.Select fg.Row, fg.Col
End Sub

Private Sub fg_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    
    ' keep current cell visible
    If chkCurrent.Value <> 0 Then
        fg.ShowCell fg.Row, fg.Col
    End If
    
End Sub

Private Sub fg_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    ' cancel if new cell is in protected area
    If chkProtect.Value <> 0 Then
        If NewRow >= 2 And NewRow <= 6 And _
           NewCol >= 2 And NewCol <= 6 Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub fg_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)

    ' cancel if editing
    If chkEditing.Value <> 0 Then
    
        ' if fg.EditWindow <> 0 then the user is editing the current cell
        If fg.EditWindow <> 0 Then
            Cancel = True
        End If
    End If
    
End Sub

Private Sub fg_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    
    ' cancel non-vertical selection
    If chkVertical.Value <> 0 Then
        If NewColSel <> fg.ColSel Then
            Cancel = True
        End If
    End If
    
    ' cancel if new selection is on protected area
    If chkProtect.Value <> 0 Then
        If fg.Row < 2 And NewRowSel < 2 Then Exit Sub
        If fg.Col < 2 And NewColSel < 2 Then Exit Sub
        If fg.Row > 6 And NewRowSel > 6 Then Exit Sub
        If fg.Col > 6 And NewColSel > 6 Then Exit Sub
        Cancel = True
    End If
    
End Sub

Private Sub Form_Load()

    ' initialize grid
    fg.Rows = 30
    fg.Editable = flexEDKbdMouse
    Dim r&, c&
    For r = fg.FixedRows To fg.Rows - 1
        For c = fg.FixedCols To fg.Cols - 1
            fg.TextMatrix(r, c) = "r" & r & "c" & c
        Next
    Next
    
    ' highlight protected range
    fg.Cell(flexcpBackColor, 2, 2, 6, 6) = RGB(200, 250, 200)

End Sub

Private Sub Form_Resize()
    
    ' resize grid
    On Error Resume Next
    fg.Move fg.Left, fg.Top, ScaleWidth - fg.Left * 2, ScaleHeight - fg.Top - fg.Left
    
    ' keep current cell visible
    If chkCurrent.Value <> 0 Then
        fg.ShowCell fg.Row, fg.Col
    End If

End Sub
