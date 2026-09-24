VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Form1 
   Caption         =   "Drag rows by the fixed cells"
   ClientHeight    =   5280
   ClientLeft      =   3000
   ClientTop       =   2070
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   7920
   Begin VB.CheckBox chkByRow 
      Caption         =   "By Row"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   3795
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   4875
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   7755
      _cx             =   13679
      _cy             =   8599
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
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkByRow_Click()

    ' set selection mode to match new state
    If chkByRow Then
        fg.SelectionMode = flexSelectionListBox
    Else
        fg.SelectionMode = flexSelectionFree
    End If
End Sub

Private Sub fg_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)

    ' allow dragging with the right mouse button on any column
    Dim r&
    r = fg.MouseRow
    If Button = vbRightButton And r >= fg.FixedRows Then
        fg.Cell(flexcpBackColor, r, fg.FixedCols, r, fg.Cols - 1) = vbRed
        r = fg.DragRow(r)
        fg.Cell(flexcpCustomFormat, r, fg.FixedCols, r, fg.Cols - 1) = False
    End If
    
End Sub

Private Sub fg_BeforeMoveRow(ByVal Row As Long, Position As Long)

    ' when selection is normal, allow default behavior
    If fg.SelectionMode <> flexSelectionListBox Then Exit Sub
    
    ' when selection is by row, drag all selected rows to new position
    With fg
        
        ' move all selected rows to bottom: this makes it easier to
        ' keep track of them as we start moving rows around
        Dim i&, r&, iSelCnt&
        iSelCnt = .SelectedRows
        For i = 0 To iSelCnt - 1
            r = .SelectedRow(0)
            .RowPosition(r) = .Rows - 1
            If r < Position Then Position = Position - 1
        Next
        
        ' move selected rows to given position
        If Position > Row Then Position = Position + 1
        For i = 0 To iSelCnt - 1
            .RowPosition(.Rows - 1) = Position
        Next
        
        ' cancel default behavior (we have handled the move)
        Position = Row
    End With
End Sub

Private Sub Form_Load()
    With fg
    
        ' initialize grid
        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeBoth
        .AllowUserFreezing = flexFreezeBoth
        .BackColorFrozen = RGB(200, 200, 255)
        .ExplorerBar = flexExMoveRows Or flexExSortShowAndMove
        
        ' fill it up with data
        Dim r%, c%
        For r = 0 To .Rows - 1
            For c = 0 To .Cols - 1
                .TextMatrix(r, c) = "r" & r & "c" & c
            Next
        Next
    
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fg.Move fg.Left, fg.Top, ScaleWidth - fg.Left * 2, ScaleHeight - fg.Top - fg.Left
End Sub
