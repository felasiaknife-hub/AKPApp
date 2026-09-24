VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "FlexGrid Database Demo"
   ClientHeight    =   4425
   ClientLeft      =   2325
   ClientTop       =   1995
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7695
   Begin MSComctlLib.ImageList imgListCategories 
      Left            =   660
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":04E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":087C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1304
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":169C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListGlyph 
      Left            =   0
      Top             =   1140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483643
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex7Ctl.VSFlexGrid fg 
      Bindings        =   "Form1.frx":1F34
      Height          =   2715
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   6015
      _cx             =   10610
      _cy             =   4789
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
      DataMode        =   1
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   1
      BackColorFrozen =   14737632
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label lbl 
      Height          =   615
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   7575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fg_AfterDataRefresh()

    ' translate product ID into product name
    Dim conn As New ADODB.Connection
    conn.Open "NorthWind"
    Dim rsList As New ADODB.Recordset
    rsList.Open "Products", conn
    fg.ColComboList(fg.ColIndex("Product")) = fg.BuildComboList(rsList, "ProductName", "ProductID")
    rsList.Close
    
    ' translate supplier ID supplier info (name, contact, phone)
    rsList.Open "Suppliers", conn
    fg.ColComboList(fg.ColIndex("Supplier")) = fg.BuildComboList(rsList, "CompanyName", "SupplierID")
    fg.ColComboList(fg.ColIndex("Country")) = fg.BuildComboList(rsList, "Country", "SupplierID")
    fg.ColComboList(fg.ColIndex("Contact")) = fg.BuildComboList(rsList, "ContactName, ContactTitle", "SupplierID")
    rsList.Close
    conn.Close
    
    ' use image list to show category pictures
    fg.ColImageList(fg.ColIndex("Cat")) = imgListCategories.hImageList
    fg.ColAlignment(fg.ColIndex("Cat")) = flexAlignCenterCenter
    
    ' do an autosize for the mapped columns
    fg.AutoSize 1, fg.Cols - 1
    
    ' use another imagelist to show db-like cursor glyph
    fg.ColWidth(0) = fg.RowHeight(0)
    fg.ColImageList(0) = imgListGlyph.hImageList
    fg.ColAlignment(0) = flexAlignCenterCenter
    fg.TextMatrix(1, 0) = 0
    
    ' freeze two columns
    fg.FrozenCols = 2
    
End Sub

Private Sub fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 3 Then Cancel = True
End Sub

Private Sub fg_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)

    ' update DB-like cursor glyph
    fg.TextMatrix(OldRow, 0) = ""
    fg.TextMatrix(NewRow, 0) = 0
    
End Sub

Private Sub fg_BeforeScrollTip(ByVal Row As Long)
    Dim Col&
    Col = fg.ColIndex("Product")
    fg.ScrollTipText = "Product: " & fg.Cell(flexcpTextDisplay, Row, Col) & " "
End Sub

Private Sub fg_DblClick()

    ' show orders for this product
    MousePointer = 11
    frmDetails.Move Left + (Width - ScaleWidth) + fg.Left + fg.ColWidth(0), _
                    Top + (Height - ScaleHeight) + fg.Top + fg.CellTop + fg.RowHeight(0)
    frmDetails.Show
    MousePointer = 0
    
End Sub

Private Sub Form_Load()
    
    ' show a little help
    lbl = "Drag the solid grid lines to freeze/thaw rows and columns." & vbCrLf & _
          "Double-click on a product to see which customers have ordered it." & vbCrLf & _
          "Category pictures come from an image list. Only product codes are stored, not their full names."
    
    ' define "NorthWind Best Sellers" SQL query (SQL created with Access)
    Dim sql$
    sql = "SELECT Products.ProductID AS Product, " & _
          "       Categories.CategoryID As Cat, " & _
          "       Suppliers.SupplierID As Supplier, " & _
          "       Suppliers.SupplierID As Country, " & _
          "       Suppliers.SupplierID As Contact, " & _
          "       Sum(CCur([Order Details].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100) AS Revenue " & _
          "FROM Suppliers INNER JOIN ((Categories INNER JOIN Products ON Categories.CategoryID = Products.CategoryID) INNER JOIN ((Customers INNER JOIN Orders ON Customers.CustomerID = Orders.CustomerID) INNER JOIN [Order Details] ON Orders.OrderID = [Order Details].OrderID) ON Products.ProductID = [Order Details].ProductID) ON Suppliers.SupplierID = Products.SupplierID " & _
          "GROUP BY Products.ProductID, Categories.CategoryID, Suppliers.SupplierID, Suppliers.SupplierID, Suppliers.SupplierID " & _
          "ORDER BY Sum(CCur([Order Details].[UnitPrice]*[Quantity]*(1-[Discount])/100)*100) DESC;"

    ' get "NorthWind Best Sellers" data
    Dim conn As New ADODB.Connection
    conn.Open "NorthWind"
    Dim rs As New ADODB.Recordset
    rs.Open sql, conn
    
    ' connect grid to data source
    fg.DataMode = flexDMBound
    Set fg.DataSource = rs
    
    ' more initialization on the AfterDataRefresh event
    fg.AllowSelection = False
    fg.Editable = flexEDKbdMouse
    fg.AllowUserResizing = flexResizeBoth
    fg.AllowUserFreezing = flexFreezeBoth
    fg.BackColorFrozen = 14737632
    fg.GridLines = flexGridFlatVert
    fg.ExtendLastCol = True
    fg.ScrollTips = True

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fg.Move fg.Left, fg.Top, ScaleWidth - 2 * fg.Left, ScaleHeight - fg.Top - fg.Left
    lbl.Width = fg.Width
End Sub
