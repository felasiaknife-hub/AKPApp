VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form Form1 
   Caption         =   "VSFlexGrid 7 Demo - Filtered ADO Data Source"
   ClientHeight    =   6285
   ClientLeft      =   2610
   ClientTop       =   2310
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   7335
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Filter"
      Height          =   675
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1275
   End
   Begin VSFlex7LCtl.VSFlexGrid fg 
      Height          =   4875
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   7095
      _cx             =   12515
      _cy             =   8599
      _ConvInfo       =   -1
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
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   615
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   5595
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_DS As FlexADO

Private Sub Command1_Click()
    
    ' clear filter, refresh data
    fg.Cell(flexcpText, 1, 0, 1, fg.Cols - 1) = ""
    fg.FlexDataSource = m_DS
    
End Sub

Private Sub fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    ' new filter, needs refresh
    If Row = 1 Then fg.FlexDataSource = m_DS
    
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    ' for this sample, we only allow editing the filter line
    If Row <> 1 Then Cancel = True
    
End Sub

Private Sub Form_Load()

    ' initialize grid (optional)
    fg.FixedCols = 0
    fg.Editable = flexEDKbdMouse
    fg.BackColorFrozen = RGB(200, 255, 200)
    
    ' create custom data source object
    Set m_DS = New FlexADO
    
    ' assign it to grid
    fg.FlexDataSource = m_DS
    fg.FrozenRows = 1
    
    ' resize based on first 6 rows (optional)
    Dim c%, r%, w%, mw%
    For c = 0 To fg.Cols - 1
        mw = 0
        For r = 0 To 5
            w = TextWidth(fg.TextMatrix(r, c))
            If w > mw Then mw = w
        Next
        fg.ColWidth(c) = mw + 100
    Next
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fg.Move fg.Left, fg.Top, ScaleWidth - fg.Left * 2, ScaleHeight - fg.Left - fg.Top
End Sub
