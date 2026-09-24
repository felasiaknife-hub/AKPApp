VERSION 5.00
Object = "{C5DE3F80-3376-11D2-BAA4-04F205C10000}#1.0#0"; "VSFLEX6D.OCX"
Begin VB.Form frmSynch 
   Caption         =   "Synchronized"
   ClientHeight    =   4860
   ClientLeft      =   1140
   ClientTop       =   1890
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   6315
   Begin VB.TextBox lblHelp 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   36
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   36
      Width           =   6132
   End
   Begin VSFlex6DAOCtl.vsFlexGrid fa 
      Height          =   1416
      Index           =   0
      Left            =   36
      TabIndex        =   0
      Top             =   756
      Width           =   3216
      _ExtentX        =   5673
      _ExtentY        =   2498
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      BackColorBkg    =   8421504
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      BackColorAlternate=   -2147483643
      SheetBorder     =   0
      Rows            =   50
      Cols            =   10
      FormatString    =   ""
      FixedRows       =   1
      FixedCols       =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      ScrollTrack     =   0   'False
      FocusRect       =   1
      HighLight       =   1
      ScrollBars      =   3
      RowHeightMin    =   0
      FillStyle       =   0
      GridLines       =   1
      GridLinesFixed  =   2
      SelectionMode   =   0
      MergeCells      =   0
      PictureType     =   0
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      VirtualData     =   -1  'True
      Editable        =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      GridLineWidth   =   1
      AutoResize      =   -1  'True
      ExtendLastCol   =   0   'False
      AutoSizeMode    =   0
      RightToLeft     =   0   'False
      MultiTotals     =   -1  'True
      OwnerDraw       =   0
      DataMode        =   0
      OleDragMode     =   0
      OleDropMode     =   0
      TabBehavior     =   0
      AllowSelection  =   -1  'True
      PicturesOver    =   0   'False
      MergeCompare    =   0
      Ellipsis        =   0
      OutlineCol      =   0
      RowHeightMax    =   0
      AutoSearch      =   0
      ExplorerBar     =   0
   End
   Begin VSFlex6DAOCtl.vsFlexGrid fa 
      Height          =   1416
      Index           =   1
      Left            =   3348
      TabIndex        =   1
      Top             =   720
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   2498
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      BackColorBkg    =   8421504
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      BackColorAlternate=   -2147483643
      SheetBorder     =   0
      Rows            =   50
      Cols            =   10
      FormatString    =   ""
      FixedRows       =   1
      FixedCols       =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      ScrollTrack     =   0   'False
      FocusRect       =   1
      HighLight       =   1
      ScrollBars      =   3
      RowHeightMin    =   0
      FillStyle       =   0
      GridLines       =   1
      GridLinesFixed  =   2
      SelectionMode   =   0
      MergeCells      =   0
      PictureType     =   0
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      VirtualData     =   -1  'True
      Editable        =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      GridLineWidth   =   1
      AutoResize      =   -1  'True
      ExtendLastCol   =   0   'False
      AutoSizeMode    =   0
      RightToLeft     =   0   'False
      MultiTotals     =   -1  'True
      OwnerDraw       =   0
      DataMode        =   0
      OleDragMode     =   0
      OleDropMode     =   0
      TabBehavior     =   0
      AllowSelection  =   -1  'True
      PicturesOver    =   0   'False
      MergeCompare    =   0
      Ellipsis        =   0
      OutlineCol      =   0
      RowHeightMax    =   0
      AutoSearch      =   0
      ExplorerBar     =   0
   End
   Begin VSFlex6DAOCtl.vsFlexGrid fa 
      Height          =   2568
      Index           =   2
      Left            =   72
      TabIndex        =   2
      Top             =   2196
      Width           =   3108
      _ExtentX        =   5482
      _ExtentY        =   4530
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      BackColorBkg    =   8421504
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      BackColorAlternate=   -2147483643
      SheetBorder     =   0
      Rows            =   50
      Cols            =   10
      FormatString    =   ""
      FixedRows       =   1
      FixedCols       =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      ScrollTrack     =   0   'False
      FocusRect       =   1
      HighLight       =   1
      ScrollBars      =   3
      RowHeightMin    =   0
      FillStyle       =   0
      GridLines       =   1
      GridLinesFixed  =   2
      SelectionMode   =   0
      MergeCells      =   0
      PictureType     =   0
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      VirtualData     =   -1  'True
      Editable        =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      GridLineWidth   =   1
      AutoResize      =   -1  'True
      ExtendLastCol   =   0   'False
      AutoSizeMode    =   0
      RightToLeft     =   0   'False
      MultiTotals     =   -1  'True
      OwnerDraw       =   0
      DataMode        =   0
      OleDragMode     =   0
      OleDropMode     =   0
      TabBehavior     =   0
      AllowSelection  =   -1  'True
      PicturesOver    =   0   'False
      MergeCompare    =   0
      Ellipsis        =   0
      OutlineCol      =   0
      RowHeightMax    =   0
      AutoSearch      =   0
      ExplorerBar     =   0
   End
   Begin VSFlex6DAOCtl.vsFlexGrid fa 
      Height          =   2604
      Index           =   3
      Left            =   3348
      TabIndex        =   3
      Top             =   2196
      Width           =   2892
      _ExtentX        =   5101
      _ExtentY        =   4593
      _ConvInfo       =   -1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      BackColorBkg    =   8421504
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      BackColorAlternate=   -2147483643
      SheetBorder     =   0
      Rows            =   50
      Cols            =   10
      FormatString    =   ""
      FixedRows       =   1
      FixedCols       =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      ScrollTrack     =   0   'False
      FocusRect       =   1
      HighLight       =   1
      ScrollBars      =   3
      RowHeightMin    =   0
      FillStyle       =   0
      GridLines       =   1
      GridLinesFixed  =   2
      SelectionMode   =   0
      MergeCells      =   0
      PictureType     =   0
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      VirtualData     =   -1  'True
      Editable        =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      GridLineWidth   =   1
      AutoResize      =   -1  'True
      ExtendLastCol   =   0   'False
      AutoSizeMode    =   0
      RightToLeft     =   0   'False
      MultiTotals     =   -1  'True
      OwnerDraw       =   0
      DataMode        =   0
      OleDragMode     =   0
      OleDropMode     =   0
      TabBehavior     =   0
      AllowSelection  =   -1  'True
      PicturesOver    =   0   'False
      MergeCompare    =   0
      Ellipsis        =   0
      OutlineCol      =   0
      RowHeightMax    =   0
      AutoSearch      =   0
      ExplorerBar     =   0
   End
End
Attribute VB_Name = "frmSynch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' dimension a variant array to hold data for
' the synchronized flex controls
Dim arr(6, 50)

Private Sub fa_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i%
    
    ' refresh other views after edits
    For i = 0 To 3
        If i <> Index Then fa(i).Refresh
    Next

End Sub

Private Sub fa_DblClick(Index As Integer)
    Dim i, j, r
    
    ' not useful, but fun
    
    With fa(Index)
    
        .ScrollBars = flexScrollBarNone
    
        r = .RowHeight(1)
        For i = r To 0 Step -4
            For j = 1 To 5
                .RowHeight(j) = i
            Next
            DoEvents
        Next
    
        For i = 0 To r Step 4
            For j = 1 To 5
                .RowHeight(j) = i
            Next
            DoEvents
        Next

        For j = 1 To 5
            .RowHeight(j) = -1
        Next
    
        .ScrollBars = flexScrollBarBoth
    
    End With
    
End Sub


Private Sub fa_RowColChange(Index As Integer)
    Dim i%
    
    ' prevent reentrancy
    ' (we'll be selecting cells, which will trigger this very event)
    Static inhere%
    inhere = False
    If inhere Then Exit Sub
    inhere = True
    
    ' synchronize selection on other views
    For i = 0 To 3
        If i <> Index Then
            fa(i).Select fa(Index).Row, fa(Index).Col
        End If
    Next
    
    ' ready for next time
    inhere = False

End Sub

Private Sub fa_Scroll(Index As Integer)

    ' synchronize scrolling
    Select Case Index
        Case 0
            fa(1).TopRow = fa(0).TopRow
            fa(2).LeftCol = fa(0).LeftCol
        Case 1
            fa(0).TopRow = fa(1).TopRow
            fa(3).LeftCol = fa(1).LeftCol
        Case 2
            fa(3).TopRow = fa(2).TopRow
            fa(0).LeftCol = fa(2).LeftCol
        Case 3
            fa(2).TopRow = fa(3).TopRow
            fa(1).LeftCol = fa(3).LeftCol
    End Select
End Sub

Private Sub fa_UserResized(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i%
    
    For i = 0 To 3
        If i <> Index Then
            If Row > -1 Then
                fa(i).RowHeight(Row) = fa(Index).RowHeight(Row)
            End If
            If Col > -1 Then
                fa(i).ColWidth(Col) = fa(Index).ColWidth(Col)
            End If
        End If
    Next
End Sub

Private Sub Form_Load()
    
    ' bind all flexes to the same array
    fa(0).BindToArray arr
    fa(1).BindToArray arr
    fa(2).BindToArray arr
    fa(3).BindToArray arr

End Sub

Private Sub Form_Resize()
    Dim w!, h!, l!, t!
    l = fa(0).Left
    
    ' adjust label position
    w = ScaleWidth - 2 * l
    lblHelp.Top = l
    If w > 0 Then lblHelp.Width = w
    
    ' adjust grid positions
    t = 2 * l + lblHelp.Height
    w = (ScaleWidth - 3 * l) / 2
    h = (ScaleHeight - l - t) / 2
    If (w > 0) And (h > 0) Then
        fa(0).Move l, t, w, h
        fa(1).Move w + 2 * l, t, w, h
        fa(2).Move l, h + l + t, w, h
        fa(3).Move w + 2 * l, h + l + t, w, h
    End If
End Sub


