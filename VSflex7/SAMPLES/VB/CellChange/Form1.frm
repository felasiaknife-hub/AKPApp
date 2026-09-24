VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsFlex7l.ocx"
Begin VB.Form Form1 
   Caption         =   "VSFlexGrid Demo: Aggregates and Conditional Formatting"
   ClientHeight    =   5715
   ClientLeft      =   1800
   ClientTop       =   1545
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4260
      Top             =   60
   End
   Begin VB.CheckBox chkAutoChange 
      Caption         =   "Make Random &Changes to the Grid"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   3855
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   3915
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   6075
      _cx             =   10716
      _cy             =   6906
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

Private Sub chkAutoChange_Click()
    
    ' enable/disable random changes to the grid
    Timer1.Enabled = chkAutoChange.Value

End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    ' can't edit frozen rows
    If Row <= fg.FrozenRows Then Cancel = True
    
End Sub

Private Sub fg_CellChanged(ByVal Row As Long, ByVal Col As Long)

    ' after a change, update totals and format cell (negative values are red)
    If Col > 0 And Row > 6 Then
    
        ' format cell:
        ' if the value is negative, make cell red and bold
        ' if > 80, make cell green and bold
        ' otherwise, reset cell format
        Dim v!
        v = fg.ValueMatrix(Row, Col)
        If v < 0 Then
            fg.Cell(flexcpForeColor, Row, Col) = RGB(150, 0, 0)
            fg.Cell(flexcpFontBold, Row, Col) = True
        ElseIf v > 80 Then
            fg.Cell(flexcpForeColor, Row, Col) = RGB(0, 150, 0)
            fg.Cell(flexcpFontBold, Row, Col) = True
        Else
            fg.Cell(flexcpCustomFormat, Row, Col) = False
        End If
        
        ' update totals
        Dim r1&, r2&
        r1 = fg.FixedRows + fg.FrozenRows
        r2 = fg.Rows - 1
        fg.TextMatrix(1, Col) = fg.Aggregate(flexSTCount, r1, Col, r2, Col)
        fg.TextMatrix(2, Col) = fg.Aggregate(flexSTSum, r1, Col, r2, Col)
        fg.TextMatrix(3, Col) = fg.Aggregate(flexSTAverage, r1, Col, r2, Col)
        fg.TextMatrix(4, Col) = fg.Aggregate(flexSTMin, r1, Col, r2, Col)
        fg.TextMatrix(5, Col) = fg.Aggregate(flexSTMax, r1, Col, r2, Col)
        fg.TextMatrix(6, Col) = fg.Aggregate(flexSTStd, r1, Col, r2, Col)
        fg.TextMatrix(7, Col) = fg.Aggregate(flexSTVar, r1, Col, r2, Col)
    End If
    
End Sub

Private Sub Form_Load()

    ' initialize grid
    With fg
        .Cols = 10
        .Rows = 30
        .ColFormat(-1) = "#,###.##"
        .TextMatrix(1, 0) = "Count"
        .TextMatrix(2, 0) = "Sum"
        .TextMatrix(3, 0) = "Average"
        .TextMatrix(4, 0) = "Min"
        .TextMatrix(5, 0) = "Max"
        .TextMatrix(6, 0) = "StdDev"
        .TextMatrix(7, 0) = "Variance"
        .FrozenRows = 7
        .BackColorFrozen = &HC0FFC0
        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeColumns
        .Select .FixedRows + .FrozenRows, 1
    End With
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fg.Move fg.Left, fg.Top, ScaleWidth - 2 * fg.Left, ScaleHeight - fg.Top - fg.Left
End Sub

Private Sub Timer1_Timer()
    
    ' randomize row and column
    Dim r&, c&
    r = Rnd * fg.Rows
    If r > fg.Rows - 1 Then r = fg.Rows - 1
    If r < fg.FixedRows + fg.FrozenRows Then r = fg.FixedRows + fg.FrozenRows
    c = Rnd * fg.Cols
    If c > fg.Cols - 1 Then c = fg.Cols - 1
    If c < fg.FixedCols Then c = fg.FixedCols
    
    ' assign random value (will update totals automatically)
    fg.TextMatrix(r, c) = Rnd * 200 - 100
    
End Sub
