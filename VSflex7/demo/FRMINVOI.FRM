VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmInvoic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VSFlexGrid Pro 6.0 - Invoice Demo"
   ClientHeight    =   6540
   ClientLeft      =   240
   ClientTop       =   1440
   ClientWidth     =   10635
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10635
   Begin VB.CommandButton cmd 
      Caption         =   "&Credit Memo"
      Height          =   495
      Index           =   1
      Left            =   9180
      TabIndex        =   1
      Top             =   1200
      Width           =   1275
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Invoice"
      Height          =   495
      Index           =   0
      Left            =   9180
      TabIndex        =   0
      Top             =   600
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5715
      Left            =   300
      ScaleHeight     =   5655
      ScaleWidth      =   8475
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Width           =   8535
      Begin VSFlex7LCtl.VSFlexGrid flexAddress 
         Height          =   1635
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   960
         Width           =   8175
         _cx             =   5080
         _cy             =   5080
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
         BackColorBkg    =   8421504
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   1
         HighLight       =   0
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
      Begin VSFlex7LCtl.VSFlexGrid flexTitle 
         Height          =   795
         Left            =   120
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   8175
         _cx             =   5080
         _cy             =   5080
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
         BackColorBkg    =   8421504
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   1
         HighLight       =   0
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
      Begin VSFlex7LCtl.VSFlexGrid flexItems 
         Height          =   2325
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3180
         Width           =   8175
         _cx             =   5080
         _cy             =   5080
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
         BackColorBkg    =   8421504
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   8
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin VSFlex7LCtl.VSFlexGrid flexTerms 
         Height          =   615
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2580
         Width           =   8175
         _cx             =   5080
         _cy             =   5080
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
         BackColorBkg    =   8421504
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   1
         HighLight       =   0
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
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   5655
      Left            =   360
      Top             =   780
      Width           =   8595
   End
End
Attribute VB_Name = "frmInvoic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' Look and Feel
Dim mDocTitle$
Dim mBackColor&

Const mForeColor& = vbWhite
Const mFontName$ = "MS Sans Serif"
Const mFontSize = "8"
Const mCellAlignment = flexAlignCenterCenter

' Options
Const mProductList$ = " VS-OCX | VSView/OCX | VSFlex VBX | VSFlexGrid Pro 7.0"
Const mShippingOptions$ = "VISA | Mastercard | AMEX | Check| NET 30| NONE"
    
Sub SetItemsGrid()
 
 Dim i%
 
 With flexItems
    
        .Redraw = False
        
        ' overall grid settings
        .FixedRows = 1
        .FixedCols = 0
        .Rows = 20
        .Cols = 6
        .ScrollBars = flexScrollBarVertical
        .GridLines = flexGridNone
        .Appearance = flexFlat
        .BorderStyle = 1
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightNever
        .Editable = True
        .SheetBorder = vbWhite
        .BackColorBkg = vbWhite
        .GridLines = flexGridFlatVert
        .FontName = "Tahoma"
        .FontSize = 8

        
        ' set column width
        .ColWidth(0) = 1065
        .ColWidth(1) = 2950
        .ColWidth(2) = 1170
        .ColWidth(3) = 1245
        .ColWidth(4) = 1170
        .ColWidth(5) = 350
        
        ' Print titles
        PrintTitle flexItems, "ITEM", 0, 0
        PrintTitle flexItems, "DESCRIPTION", 0, 1
        PrintTitle flexItems, "QTY", 0, 2
        PrintTitle flexItems, "RATE", 0, 3
        PrintTitle flexItems, "AMOUNT", 0, 4
        PrintTitle flexItems, "Tax", 0, 5
        
        ' apply common formatting
        .Select 0, 0, 0, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
 
        .Redraw = True
    End With
End Sub

Sub SetTermsGrid()

   With flexTerms
    
        .Redraw = False
        
        ' overall grid settings
        .FixedRows = 0
        .FixedCols = 0
        .Rows = 4
        .Cols = 12
        .ScrollBars = flexScrollBarNone
        .GridLines = flexGridNone
        .Appearance = flexFlat
        .BorderStyle = 0
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightNever
        .Editable = True
        .FillStyle = flexFillRepeat
        
        .RowHeight(0) = 10
        .RowHeight(3) = 10
        
        .SheetBorder = vbWhite
        .BackColorBkg = vbWhite

        
        ' set column width
        .ColWidth(0) = 500
        .ColWidth(1) = 1200
        .ColWidth(2) = 10
        .ColWidth(3) = 1600
        .ColWidth(4) = 10
        .ColWidth(5) = 1200
        .ColWidth(6) = 10
        .ColWidth(7) = 1200
        .ColWidth(8) = 10
        .ColWidth(9) = 1200
        .ColWidth(10) = 10
        .ColWidth(11) = 1200

        ' PO NO.
        PrintTitle flexTerms, "PO NO.", 1, 1
        PrintTitle flexTerms, "TERMS", 1, 3
        PrintTitle flexTerms, "REP", 1, 5
        PrintTitle flexTerms, "SHIP DATE", 1, 7
        PrintTitle flexTerms, "SHIP VIA", 1, 9
        PrintTitle flexTerms, "FOB", 1, 11
        
        .Select 1, 1, 2, 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select 1, 3, 2, 3
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select 1, 5, 2, 5
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select 1, 7, 2, 7
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select 1, 9, 2, 9
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select 1, 11, 2, 11
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        
        .Redraw = True
    End With
End Sub

Sub PrintTitle(ByRef flex As VSFlexGrid, mText$, mRow&, mCol&)

    With flex
        .TextMatrix(mRow, mCol) = mText
        .Select mRow, mCol
        .CellBackColor = mBackColor
        .CellForeColor = mForeColor
        .CellFontName = mFontName
        .CellFontSize = mFontSize
        .CellAlignment = mCellAlignment
    End With
    
End Sub

Sub SetTitleGrid()
    
    ' FlexTitle
    With flexTitle
    
        .Redraw = False
        
        ' overall grid settings
        .FixedRows = 0
        .FixedCols = 0
        .Rows = 4
        .Cols = 5
        .ScrollBars = flexScrollBarNone
        .GridColor = .BackColor
        .Appearance = flexFlat
        .BorderStyle = 0
        .FocusRect = flexFocusNone
        .Editable = True
        .RowHeight(3) = 10
        .SheetBorder = vbWhite
        .BackColorBkg = vbWhite
        .FillStyle = flexFillRepeat
        .MergeCells = flexMergeRestrictColumns
        
        ' set columns width
        .ColWidth(0) = 6000
        .ColWidth(1) = 975
        .ColWidth(2) = 50
        .ColWidth(3) = 1155
        
        ' Place title
        .MergeCol(0) = True
        .Cell(flexcpText, 0, 0, 1, 0) = mDocTitle
        .Cell(flexcpFontName, 0, 0, 1, 0) = "Haettenschweiler"
        .Cell(flexcpFontSize, 0, 0, 1, 0) = "24"
        .Cell(flexcpForeColor, 0, 0, 1, 0) = mBackColor
        
        ' invoice No.
        PrintTitle flexTitle, "DATE", 1, 1
        PrintTitle flexTitle, "INVOICE NO.", 1, 3
        
        ' draw borders
        .Select 1, 3, 2, 3
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select 1, 1, 2, 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        
        .Redraw = True
        
    End With
    
End Sub


Private Sub cmd_Click(Index%)

    Select Case Index
        Case 0
            mDocTitle = "INVOICE"
            mBackColor = &H800000  ' Blue
        Case 1
            mDocTitle = "CREDIT MEMO"
            mBackColor = &HC0&            ' Red
    End Select
    
    ' draw grids
    SetTitleGrid
    SetAddressGrid
    SetTermsGrid
    SetItemsGrid
    
End Sub

Private Sub flexItems_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    With flexItems
    
        ' place a number on the price
        If .ValueMatrix(Row, 2) = 0 Then .TextMatrix(Row, 2) = 1
    
        Select Case .EditText
            Case " VSVBX "
                .TextMatrix(Row, 1) = "VideoSoft VSVBX"
                .TextMatrix(Row, 3) = Format("45", "$#,###,##0.00")
                .TextMatrix(Row, 4) = Format(.ValueMatrix(Row, 2) * .ValueMatrix(Row, 3), "$#,###,##0.00")
            Case " VSOCX "
                .TextMatrix(Row, 1) = "VideoSoft VS-OCX Version 5.0"
                .TextMatrix(Row, 3) = "99.00"
                .TextMatrix(Row, 4) = Format(.ValueMatrix(Row, 2) * .ValueMatrix(Row, 3), "$#,###,##0.00")
            Case Else
                Exit Sub
        End Select
    End With

End Sub

Private Sub flexItems_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Select Case Col
        Case 0
            flexItems.ComboList = mProductList
        Case 1 ' Description
            Cancel = True
        Case Else
            flexItems.ComboList = ""
    End Select

End Sub

Private Sub flexItems_Click()
    flexItems.EditCell
End Sub


Sub SetAddressGrid()

    With flexAddress
    
        .Redraw = False
        
        ' overall grid settings
        .FixedRows = 0
        .FixedCols = 0
        .Rows = 4
        .Cols = 5
        .ScrollBars = flexScrollBarNone
        .GridColor = .BackColor
        .Appearance = flexFlat
        .BorderStyle = 0
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightNever
        .Editable = True
        .RowHeight(0) = 50
        .RowHeight(2) = 1170
        .RowHeight(3) = 10
        .SheetBorder = vbWhite
        .BackColorBkg = vbWhite
        .FillStyle = flexFillRepeat
        
        ' set column width
        .ColWidth(0) = 100
        .ColWidth(2) = 2880
        .ColWidth(3) = 2600
        .ColWidth(1) = 2600
        
        ' Print Fields
        PrintTitle flexAddress, "    BILL TO", 1, 1
        .CellAlignment = flexAlignLeftCenter
        PrintTitle flexAddress, "    SHIP TO", 1, 3
        .CellAlignment = flexAlignLeftCenter
        
        ' draw boxes
        .Select 1, 1, 2, 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select 1, 3, 2, 3
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        
        .Redraw = True
        
    End With
End Sub

