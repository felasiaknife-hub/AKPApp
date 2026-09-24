VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Form1 
   Caption         =   "VSFlexGrid 7 - Custom Data Source Demo"
   ClientHeight    =   5715
   ClientLeft      =   2310
   ClientTop       =   1800
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   6585
   Begin VSFlex7LCtl.VSFlexGrid fgLinks 
      Height          =   2340
      Left            =   225
      TabIndex        =   1
      Top             =   3150
      Width           =   5640
      _cx             =   9948
      _cy             =   4128
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
   Begin VSFlex7LCtl.VSFlexGrid fgNodes 
      Height          =   2415
      Left            =   75
      TabIndex        =   0
      Top             =   375
      Width           =   5790
      _cx             =   10213
      _cy             =   4260
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Links"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   150
      TabIndex        =   3
      Top             =   2850
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nodes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   75
      TabIndex        =   2
      Top             =   150
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fgLinks_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    ' after changing link connectivity, show new length
    If Col = 2 Or Col = 3 Then fgLinks.Cell(flexcpRefresh, Row, 4) = True

End Sub

Private Sub fgLinks_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    ' cannot edit link length
    If Col = 4 Then Cancel = True
    
End Sub

Private Sub fgNodes_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    ' after moving nodes, update link lengths
    ' note that setting the Redraw property automatically refreshes the control
    If Col > 1 Then fgLinks.Redraw = flexRDBuffered
    
End Sub

Private Sub Form_Load()

    ' initialize data
    InitData
    
    ' create data acess objects
    Dim dsNodes As New FlexDataSourceNodes
    Dim dsLinks As New FlexDataSourceLinks
    
    ' display node data
    InitGrid fgNodes
    With fgNodes
        .FlexDataSource = dsNodes
        .ColFormat(2) = "#.00"
        .ColFormat(3) = "#.00"
        .ColFormat(4) = "#.00"
    End With
    
    ' display link data
    InitGrid fgLinks
    With fgLinks
        .FlexDataSource = dsLinks
        .ColFormat(4) = "#.00"
        .ColFormat(5) = "#.00"
        .ColFormat(6) = "#.00"
        .ColFormat(7) = "#.00"
        .ColFormat(8) = "#.00"
    End With
    
End Sub

Private Sub Form_Resize()

    With fgNodes
        .Move .Left, .Top, ScaleWidth - 2 * .Left, ScaleHeight / 2 - .Top - .Left
        fgLinks.Move .Left, 2 * .Top + .Height, .Width, .Height
        Label1.Move .Left, .Top - Label1.Height * 1.1
        Label2.Move .Left, fgLinks.Top - Label2.Height * 1.1
    End With

End Sub

Sub InitGrid(f As VSFlexGrid)
    Dim i%
    With f
        .Editable = flexEDKbdMouse
        .FrozenCols = 1
        .ColWidth(0) = 700
        .ColWidth(1) = 1200
        .AllowUserResizing = flexResizeBothUniform
        .AllowUserFreezing = flexFreezeBoth
        .BackColorFrozen = &HFFC0C0
        For i = 2 To .Cols - 1
            .ColAlignment(i) = flexAlignRightTop
        Next
    End With
End Sub
