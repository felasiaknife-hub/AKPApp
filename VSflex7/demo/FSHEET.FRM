VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form fSheet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Balance Sheet"
   ClientHeight    =   7620
   ClientLeft      =   495
   ClientTop       =   840
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkWallpaper 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Draft Wallpaper"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5460
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VSFlex7LCtl.VSFlexGrid flex 
      Height          =   6015
      Left            =   420
      TabIndex        =   0
      Top             =   1320
      Width           =   9750
      _cx             =   5080
      _cy             =   5080
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   15
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
      OutlineBar      =   1
      OutlineCol      =   3
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
      WallPaperAlignment=   10
   End
   Begin VB.Image imgWallpaper 
      Height          =   2730
      Left            =   9180
      Picture         =   "fsheet.frx":0000
      Top             =   6780
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   1
      Left            =   420
      Shape           =   3  'Circle
      Top             =   780
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   540
      X2              =   3300
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Image imgTitle 
      Height          =   870
      Left            =   600
      Picture         =   "fsheet.frx":1D7F
      Top             =   180
      Width           =   2370
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   5895
      Index           =   0
      Left            =   660
      Top             =   1560
      Width           =   9615
   End
   Begin VB.Image Image4 
      Height          =   165
      Left            =   9600
      Picture         =   "fsheet.frx":3800
      Top             =   180
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgClear 
      Height          =   375
      Left            =   8580
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   8340
      Picture         =   "fsheet.frx":39A2
      Top             =   120
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image Image3 
      Height          =   225
      Left            =   8100
      Picture         =   "fsheet.frx":3CD0
      Top             =   120
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   9180
      Picture         =   "fsheet.frx":3F6A
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "fSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim t#              ' column # for the amounts
Sub DoGroup(Row, lvl)
    
    ' set the row as a group
    flex.IsSubtotal(Row) = True
    
    ' set the indentation level of the group
    flex.RowOutlineLevel(Row) = lvl
    
    
    If lvl = 0 Then Exit Sub
    
    If lvl = 1 Then
        flex.Cell(flexcpForeColor, Row, t) = vbWhite
        flex.Cell(flexcpBackColor, Row, t) = 1
    End If
    
    flex.Cell(flexcpFontBold, Row, t) = True
    
    
    
End Sub


Private Sub chkWallpaper_Click()

    If chkWallpaper Then
        flex.WallPaper = imgWallpaper
    Else
        flex.WallPaper = Nothing
    End If

End Sub

Private Sub flex_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)

    ' show popup menu on the date
    If Row = 3 And col = 8 Then
        flex.ComboList = "..."
    
    ' other cells are not editable
    Else
        flex.ComboList = ""
        Cancel = True
    End If
    

End Sub

Private Sub flex_BeforeScrollTip(ByVal Row As Long)
    
    

    If Row < 3 Or flex.Cell(flexcpTextDisplay, Row, 3) = "" Then Exit Sub
    flex.ScrollTipText = " " & Trim$(flex.Cell(flexcpTextDisplay, Row, 3)) & " "

End Sub

Private Sub flex_CellButtonClick(ByVal Row As Long, ByVal col As Long)

    ' show a message when user clicks on the ellipsis button
    MsgBox "The next update is scheduled for March 12, 2002", vbInformation
End Sub

Private Sub flex_DblClick()
    
    ' allow collapsing/expanding with double-clicks
    If flex.IsCollapsed(flex.Row) = flexOutlineCollapsed Then
        flex.IsCollapsed(flex.Row) = flexOutlineExpanded
    Else
        flex.IsCollapsed(flex.Row) = flexOutlineCollapsed
    End If
End Sub

Private Sub flex_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i#
    Static r&, c&

    ' avoid extra work
    If flex.MouseRow = r Then Exit Sub
    
    ' MouseRow returns -1 when not over sheet
    If flex.MouseRow = -1 Then Exit Sub
    r = flex.MouseRow

    ' move selection (even with no click)
    flex.Select r, 8
    
    ' remove cursor image
    flex.Cell(flexcpPicture, 4, 7, flex.Rows - 1) = imgClear
    flex.Cell(flexcpPicture, 4, 9, flex.Rows - 1) = imgClear

    ' show cursor image if there is some text in column 5
    If r < 4 Or flex.Cell(flexcpText, r, 8) = "" Then Exit Sub
    flex.Cell(flexcpPicture, r, 7) = Image3.Picture
    flex.Cell(flexcpPicture, r, 9) = Image1.Picture

End Sub




Private Sub Form_Load()

    Dim i%, sp$, mCol%
    
    ' set variables
    mCol = 3                ' tree column
    t = 8                   ' amount column
    sp = Space$(5)          ' outline indentention
    
    With flex
    
        ' set column width
        For i = 0 To 5
            .ColWidth(i) = 200
        Next i
        .ColWidth(mCol) = 800   ' more space for the tree column
        
        ' set global settings
        .GridLines = flexGridNone
        .ScrollTrack = False
        .ScrollTips = True
        .Editable = True
        .BackColorBkg = vbWhite
        .SheetBorder = vbWhite
        .Cell(flexcpFontBold, 4, t, .Rows - 1, t) = False   ' set font for amounts
        
        ' set title
        .Cell(flexcpText, 1, 3) = "The Company Inc."
        .Cell(flexcpFontSize, 1, 3) = 12
        .Cell(flexcpForeColor, 1, 3) = vbBlue
        .RowHeight(1) = 300
        
        ' set subtitle
        .Cell(flexcpText, 2, 3) = "A ficticious company"
        .Cell(flexcpFontBold, 2, 3) = False
        .Cell(flexcpFontItalic, 2, 3) = True
        .Cell(flexcpText, 3, t) = "As of " & Format$(Now, "mmm dd, yyyy") & "   "
        .Cell(flexcpForeColor, 3, 3) = vbBlue
        .Cell(flexcpPicture, 3, t - 1) = Image4.Picture
        .Cell(flexcpPicture, 3, t + 1) = Image4.Picture
        
        '------------------------------------------------------------
        ' create outline structure
        '------------------------------------------------------------
        DoGroup 3, 0
        DoGroup 4, 1
        DoGroup 5, 2
        DoGroup 6, 3
        DoGroup 13, 3
        DoGroup 17, 3
        DoGroup 26, 2
        DoGroup 27, 3
        DoGroup 33, 1
        DoGroup 34, 2
        DoGroup 35, 3
        DoGroup 40, 2
        
        '------------------------------------------------------------
        ' create outline nodes and data
        '------------------------------------------------------------
        .Cell(flexcpText, 3, mCol) = "BALANCE SHEET"
        .Cell(flexcpText, 4, mCol) = "ASSETS"
        .Cell(flexcpText, 4, t) = "3,972,500.00"
        
        .Cell(flexcpText, 5, mCol) = "Current Assets"
        .Cell(flexcpText, 5, t) = "1,232,500.00"
        
        .Cell(flexcpText, 6, mCol) = "Checking/Savings"
        .Cell(flexcpText, 6, t) = "340,070.00"
        .Cell(flexcpText, 7, mCol) = sp & "Bank - Checking"
        .Cell(flexcpText, 7, t) = "32,500.00"
        .Cell(flexcpText, 8, mCol) = sp & "Bank - Money Market"
        .Cell(flexcpText, 8, t) = "140,070.00"
        .Cell(flexcpText, 9, mCol) = sp & "Bank VideoSoft -  Money Market"
        .Cell(flexcpText, 9, t) = "94,080.00"
        .Cell(flexcpText, 10, mCol) = sp & "Bank VideoSoft - Savings"
        .Cell(flexcpText, 10, t) = "73,420.00"
        .Cell(flexcpText, 11, mCol) = "Total Checking/Savings"
        .Cell(flexcpText, 11, t) = "340,070.00"
        
        .Cell(flexcpText, 13, mCol) = "Accounts Receivable"
        .Cell(flexcpText, 13, t) = "280,500.00"
        .Cell(flexcpText, 14, mCol) = sp & sp & "Accounts Receivable"
        .Cell(flexcpText, 14, t) = "280,500.00"
        .Cell(flexcpText, 15, mCol) = "Total Accounts Receivable"
        .Cell(flexcpText, 14, t) = "280,500.00"
        .Cell(flexcpText, 15, t) = "280,500.00"
        
        .Cell(flexcpText, 17, mCol) = "Other Current Assets"
        .Cell(flexcpText, 17, t) = "611,930.00"
        .Cell(flexcpText, 18, mCol) = sp & sp & "Other Receivables"
        .Cell(flexcpText, 18, t) = "32,500.00"
        .Cell(flexcpText, 19, mCol) = sp & sp & "Short Term Investments"
        .Cell(flexcpText, 19, t) = "432,100.00"
        .Cell(flexcpText, 20, mCol) = sp & sp & "Inventory"
        .Cell(flexcpText, 20, t) = "15,440.00"
        .Cell(flexcpText, 21, mCol) = sp & sp & "Pre-paid Expenses"
        .Cell(flexcpText, 21, t) = "131,890.00"
        .Cell(flexcpText, 22, mCol) = "Total Other Current Assets"
        .Cell(flexcpText, 22, t) = "611,930.00"
        
        .Cell(flexcpText, 24, mCol) = "Total Current Assets"
        .Cell(flexcpText, 24, t) = "1,232,500.00"
        .Cell(flexcpBackColor, 24, t) = &H8000000F
        
        .Cell(flexcpText, 26, mCol) = "Fixed Assets"
        .Cell(flexcpText, 26, t) = "2,740,000.00"
        .Cell(flexcpText, 27, mCol) = "Fixed Assets"
        .Cell(flexcpText, 27, t) = "2,740,000.00"
        .Cell(flexcpText, 28, mCol) = sp & sp & "Accumulated Depreciation"
        .Cell(flexcpText, 28, t) = "1,240,000.00"
        .Cell(flexcpText, 29, mCol) = sp & sp & "Equipment"
        .Cell(flexcpText, 29, t) = "1,500,000.00"
        .Cell(flexcpText, 30, mCol) = "Total Fixed Assets"
        .Cell(flexcpText, 30, t) = "2,740,000.00"
        
        .Cell(flexcpText, 33, mCol) = "LIABILITIES & EQUITY"
        .Cell(flexcpText, 33, t) = "3,972,500.00"
        .Cell(flexcpText, 34, mCol) = "Liabilities"
        .Cell(flexcpText, 34, t) = "235,400.00"
        .Cell(flexcpText, 35, mCol) = "Current Liabilities"
        .Cell(flexcpText, 35, t) = "235,400.00"
        .Cell(flexcpText, 36, mCol) = sp & sp & "Accounts Payable"
        .Cell(flexcpText, 36, t) = "64,250.00"
        .Cell(flexcpText, 37, mCol) = sp & sp & "Payroll Liabilities"
        .Cell(flexcpText, 37, t) = "171,150.00"
        .Cell(flexcpText, 38, mCol) = "Total Current Liabilities"
        .Cell(flexcpText, 38, t) = "235,400.00"
        
        .Cell(flexcpText, 40, mCol) = "Equity"
        .Cell(flexcpText, 40, t) = "3,737,100.00"
        .Cell(flexcpText, 41, mCol) = sp & sp & "Retained Earnings"
        .Cell(flexcpText, 41, t) = "570,000.00"
        .Cell(flexcpText, 42, mCol) = sp & sp & "Net Income"
        .Cell(flexcpText, 42, t) = "1,750,000.00"
        .Cell(flexcpText, 43, mCol) = sp & sp & "Common Stock"
        .Cell(flexcpText, 43, t) = "640,000.00"
        .Cell(flexcpText, 44, mCol) = sp & sp & "Paid-In Capital"
        .Cell(flexcpText, 44, t) = "777,100.00"
        .Cell(flexcpText, 45, mCol) = "Total Equity"
        .Cell(flexcpText, 45, t) = "3,737,100.00"
        
        ' subtotals cell lines
        .Row = 11: .col = t
        .CellBorder vbBlack, 0, 1, 0, 0, 0, 0
        .Row = 15: .col = t
        .CellBorder vbBlack, 0, 1, 0, 0, 0, 0
        .Row = 22: .col = t
        .CellBorder vbBlack, 0, 1, 0, 0, 0, 0
        .Row = 24: .col = t
        .CellBorder vbBlack, 0, 1, 0, 0, 0, 0
        .Row = 30: .col = t
        .CellBorder vbBlack, 0, 1, 0, 0, 0, 0
        .Row = 38: .col = t
        .CellBorder vbBlack, 0, 1, 0, 0, 0, 0
        .Row = 45: .col = t
        .CellBorder vbBlack, 0, 1, 0, 0, 0, 0
        
        ' make it look good
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize mCol
        .AutoSize t
        .MergeCells = flexMergeSpill
        .ColWidth(t - 1) = 300
        .ColWidth(t + 1) = 300
    End With
End Sub


