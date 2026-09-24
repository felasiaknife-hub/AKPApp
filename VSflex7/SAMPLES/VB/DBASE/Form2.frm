VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Begin VB.Form frmDetails 
   BorderStyle     =   0  'None
   Caption         =   "Orders for the current product"
   ClientHeight    =   2355
   ClientLeft      =   3105
   ClientTop       =   3405
   ClientWidth     =   5430
   LinkTopic       =   "Form2"
   ScaleHeight     =   2355
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VSFlex7Ctl.VSFlexGrid fg 
      Height          =   2340
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      _cx             =   9551
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
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
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
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fg_AfterDataRefresh()
    
    ' use another imagelist to show db-like cursor glyph
    fg.ColWidth(0) = fg.RowHeight(0)
    fg.ColImageList(0) = frmMain!imgListGlyph.hImageList
    fg.ColAlignment(0) = flexAlignCenterCenter
    fg.TextMatrix(1, 0) = 0

    ' freeze one column
    fg.FrozenCols = 1
    
    ' size form so it fits just right
    If fg.RowHeight(0) * fg.Rows < fg.ClientHeight Then
        fg.Height = fg.RowHeight(0) * fg.Rows + (fg.Height - fg.ClientHeight)
        Height = fg.Height + fg.Top * 2
    End If
    
End Sub

Private Sub fg_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)

    ' update DB-like cursor glyph
    fg.TextMatrix(OldRow, 0) = ""
    fg.TextMatrix(NewRow, 0) = 0
    
End Sub

Private Sub fg_BeforeScrollTip(ByVal Row As Long)
    Dim Col&
    Col = fg.ColIndex("Customer")
    fg.ScrollTipText = " Customer: " & fg.TextMatrix(Row, Col) & " "
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    
    ' get current product id
    Dim prodID&
    With frmMain.fg
        prodID = .TextMatrix(.Row, .ColIndex("Product"))
    End With
    
    ' define "NorthWind Best Sellers Details" SQL query (SQL created with Access)
    Dim sql$
    sql = "SELECT Customers.CompanyName As Customer, Orders.OrderDate As [Date], CCur([Order Details].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100 AS [Value] " & _
          "FROM Customers INNER JOIN (Orders INNER JOIN [Order Details] ON Orders.OrderID = [Order Details].OrderID) ON Customers.CustomerID = Orders.CustomerID " & _
          "WHERE ((([Order Details].ProductID) = " & prodID & "))" & _
          "ORDER BY CCur([Order Details].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100 DESC;"

    ' get "NorthWind Best Sellers Details" data
    Dim conn As New ADODB.Connection
    conn.Open "NorthWind"
    Dim rs As New ADODB.Recordset
    rs.Open sql, conn

    ' connect grid to data source
    Set fg.DataSource = rs
    
    ' more initialization on the AfterDataRefresh event
    fg.AllowSelection = False
    fg.AllowUserResizing = flexResizeBoth
    fg.AllowUserFreezing = flexFreezeBoth
    fg.BackColorFrozen = 14737632
    fg.GridLines = flexGridFlatVert
    fg.ExtendLastCol = True
    fg.ScrollTips = True
    
End Sub
