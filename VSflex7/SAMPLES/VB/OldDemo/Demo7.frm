VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "VSFLEX7L.OCX"
Begin VB.Form frmMain 
   Caption         =   "VideoSoft VSFLEX6 Demo"
   ClientHeight    =   4950
   ClientLeft      =   1785
   ClientTop       =   1965
   ClientWidth     =   7875
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   7875
   Begin VB.CommandButton Command1 
      Caption         =   "HTML"
      Height          =   375
      Left            =   1980
      TabIndex        =   4
      Top             =   270
      Width           =   795
   End
   Begin VB.TextBox lblHelp 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C00000&
      Height          =   600
      Left            =   2940
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   90
      Width           =   3705
   End
   Begin VB.CommandButton btnGo 
      Caption         =   "Show Me"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   684
      Left            =   84
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   84
      Width           =   2730
   End
   Begin VSFlex7LCtl.VSFlexGrid lstMenu 
      Height          =   3870
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2715
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0   'False
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
   End
   Begin VSFlex7LCtl.VSFlexGrid fa 
      Height          =   3840
      Left            =   2880
      TabIndex        =   1
      Top             =   840
      Width           =   4335
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   12
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
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
      Editable        =   0   'False
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
   End
   Begin VB.Image imgFolder 
      Height          =   225
      Left            =   7470
      Picture         =   "Demo7.frx":0000
      Top             =   1230
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgItem 
      Height          =   225
      Left            =   7515
      Picture         =   "Demo7.frx":00FA
      Top             =   975
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image imgCell 
      Height          =   150
      Left            =   7350
      Picture         =   "Demo7.frx":01F4
      Top             =   2340
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image imgNope 
      Height          =   480
      Left            =   7335
      Picture         =   "Demo7.frx":02EE
      Top             =   1590
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgDrag 
      Height          =   480
      Left            =   6975
      Picture         =   "Demo7.frx":05F8
      Top             =   255
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type DEMOINFO_TYPE
    name As String
    help As String
End Type

Dim Demo() As DEMOINFO_TYPE
Dim CurrentDemo As String

Const slFileNames = "File Name|Msv2cmvr.Ax|Unaxa.Exe|Ieshwiz.Exe|Ie4tour.Dll|Webvw.Dll|Javacypt.Dll|Javaprxy.Dll|Javart.Dll"
Const slFileSizes = "File Size|123123|234234|34534|345455|654566|789997|234444|23444|3422|232344|24443|663453|534535"
Const slFileTypes = "Type|System|User|Application|Library|Log|Trash"
Const slFileDates = "Date|3/4/98|2/5/97|6/8/97|2/4/96|2/6/98|9/2/96|1/5/94|1/4/96"
Const slFileHidden = "Hidden|0|-1"
Const slLastNames = "Last Name|Mindelis|Blackmore|Howe|Rabin|Barre|Satriani|Van Halen|May|Solari"
Const slFirstNames = "First Name|Alex|Bonnie|Charles|Daniel|Edward|Fred|Gerry|Howard|Irene|Jackie|Katherine|Louise|Martha|Nora"
Const slPhones = "Phone||||(415) 555-1212|(415) 555-2121|(415) 666-1212|(415) 555-4323|(510) 234-4443|(212) 234-4323|(212) 434-8909"
Const slCards = "Card||||1111 2222 3333 4444|2222 3333 4444 5555|3333 4444 5555 6666|4444 5555 6666 7777|5555 6666 7777 8888"
Const slComboNames = "#111*1&16761024;Nuno" & vbTab & "Mindelis" & vbTab & "Broadway, 1232" & vbTab & "972-2321" & _
                      "|#222;Richie" & vbTab & "Blackmore" & vbTab & "Spruce, 500-T" & vbTab & "342-2321" & _
                      "|#333;Steve" & vbTab & "Howe" & vbTab & "Lombard, 543" & vbTab & "011 55 023/955-4561" & _
                      "|#444;Trevor" & vbTab & "Rabin" & vbTab & "Main, 1232" & vbTab & "972-2321" & _
                      "|#555;Martin" & vbTab & "Barre" & vbTab & "Marine, 1137" & vbTab & "345-2342" & _
                      "|#666;Joe" & vbTab & "Satriani" & vbTab & "Los Angeles, 500-T" & vbTab & "342-2321" & _
                      "|#777;Eddie" & vbTab & "Van Halen" & vbTab & "Madera, 543" & vbTab & "011 55 023/955-4561" & _
                      "|#888;Brian" & vbTab & "May" & vbTab & "Bullock, 500-T" & vbTab & "342-2321" & _
                      "|#999;Frank" & vbTab & "Solari" & vbTab & "Geary, 543" & vbTab & "011 55 023/955-4561"
Const slComboPhones = "|#1*1&12648384;Main Office" & vbTab & "(415) 555-1212" & _
                      "|#2;East Coast Office" & vbTab & "(212) 444-1212" & _
                      "|#3;Main FAX" & vbTab & "(415) 333-1212" & _
                      "|#4;Order Desk" & vbTab & "(800) 222-1212" & _
                      "|#5;Help Desk" & vbTab & "(800) 111-1212"

Const slProduct = "Product|Flutes|Saxophones|Drums|Guitars|Trombones|Keyboards|Microphones"
Const slAssociate = "Associate|John|Paul|Mike|Paula|Sylvia|Donna"
Const slRegion = "Region|North|South|East|West"
Const slSales = "Sales|14323|2532|45342|43432|75877|4232|4543"

' API declares
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long




Sub DemoAutoSearch()
    
    ' reset the control
    SetDefaults fa
    
    ' set the properties we want
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.ExplorerBar = flexExSortShowAndMove
    
    ' fill the control with data
    fa.Rows = 50
    fa.FixedRows = 1
    fa.Cols = 4
    fa.ColFormat(1) = "#,###"
    fa.ColDataType(3) = flexDTDate
    fa.ColFormat(3) = "Medium Date"
    FillColumn fa, 0, slFileNames
    FillColumn fa, 1, slFileSizes
    FillColumn fa, 2, slFileTypes
    FillColumn fa, 3, slFileDates
    
    ' let user play with it:
    ' - search by typing on any column
    '   partial matches will be highlighted and the cursor will move to show the selection
    fa.AutoSearch = flexSearchFromCursor

End Sub

Sub DemoCombo()
    
    ' reset the control
    SetDefaults fa
    
    ' set the properties we want
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.Editable = True
    
    ' fill the control with data
    fa.Rows = 50
    fa.FixedRows = 1
    fa.Cols = 3
    
    FillColumn fa, 0, slLastNames
    fa.ColComboList(0) = slComboNames
    
    FillColumn fa, 1, slPhones
    fa.Cell(flexcpText, 0, 1) = "Office Phone"
    fa.ColEditMask(1) = "(###) 000-0000"
    fa.ColComboList(1) = slComboPhones
    
    FillColumn fa, 2, slPhones
    fa.Cell(flexcpText, 0, 2) = "Home Phone"
    fa.ColEditMask(2) = "(###) 000-0000"
    fa.ColComboList(2) = slComboPhones
    
    ' make all columns same width
    fa.ColWidth(-1) = fa.ClientWidth / fa.Cols
    
    ' let user play with it:
    
End Sub

Sub DemoPopup()
    
    ' reset the control
    SetDefaults fa
    
    ' set the properties we want
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.ExplorerBar = flexExSortShowAndMove
    fa.Editable = True
    
    ' fill the control with data
    fa.Rows = 50
    fa.FixedRows = 1
    fa.Cols = 6
    
    FillColumn fa, 0, slFileNames
    fa.ColComboList(0) = "..."  ' file picker popup
    
    FillColumn fa, 1, slFileSizes
    fa.ColFormat(1) = "#,###"
    
    FillColumn fa, 2, slFileHidden
    fa.ColDataType(2) = flexDTBoolean
    
    FillColumn fa, 3, slFileDates
    fa.ColDataType(3) = flexDTDate
    fa.ColFormat(3) = "Short Date"
    fa.Cell(flexcpText, 0, 3) = "Short Date"
    fa.ColComboList(3) = "..."  ' date picker popup
    
    FillColumn fa, 4, slFileDates
    fa.ColDataType(4) = flexDTDate
    fa.ColFormat(4) = "Medium Date"
    fa.Cell(flexcpText, 0, 4) = "Medium Date"
    fa.ColComboList(4) = "..."  ' date picker popup
    
    FillColumn fa, 5, slFileDates
    fa.ColDataType(5) = flexDTDate
    fa.ColFormat(5) = "yyyy.mm.dd"
    fa.Cell(flexcpText, 0, 5) = "Custom Date"
    fa.ColComboList(5) = "..."  ' date picker popup
    
    ' do an autosize
    fa.AutoSize 0, fa.Cols - 1, , 300

End Sub

Sub DemoFormat()
    
    ' reset the control
    SetDefaults fa
    
    ' set the properties we want
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.ExplorerBar = flexExSortShowAndMove
    fa.Editable = True
    
    ' fill the control with data
    fa.Rows = 50
    fa.FixedRows = 1
    fa.Cols = 6
    
    FillColumn fa, 0, slFileNames
    
    FillColumn fa, 1, slFileSizes
    fa.ColFormat(1) = "#,###"
    
    FillColumn fa, 2, slFileHidden
    fa.ColDataType(2) = flexDTBoolean
    
    FillColumn fa, 3, slFileDates
    fa.ColDataType(3) = flexDTDate
    fa.ColFormat(3) = "Short Date"
    fa.Cell(flexcpText, 0, 3) = "Short Date"
    
    FillColumn fa, 4, slFileDates
    fa.ColDataType(4) = flexDTDate
    fa.ColFormat(4) = "Medium Date"
    fa.Cell(flexcpText, 0, 4) = "Medium Date"
    
    FillColumn fa, 5, slFileDates
    fa.ColDataType(5) = flexDTDate
    fa.ColFormat(5) = "yyyy.mm.dd"
    fa.Cell(flexcpText, 0, 5) = "Custom Date"
    
    ' do an autosize
    fa.AutoSize 0, fa.Cols - 1, , 300

End Sub


Sub DemoBindToArrayADO()
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    ' open the connection
    Set conn = New ADODB.Connection
    On Error Resume Next
    conn.Open "NorthWind"
    If Err Then
        MsgBox "Sorry, but this demo requires an ADO datasource named 'NorthWind'. You don't seem to have one set up."
        Exit Sub
    End If
    
    ' get some data into a variant array
    ' note that the array is dimensioned as a static, otherwise it would
    ' be nuked at the end of this routine and we would have nothing to bind to
    Set rs = conn.Execute("SELECT * FROM Customers")
    Static arr As Variant
    arr = rs.GetRows
    
    ' reset the control
    SetDefaults fa
    
    ' set the properties we want
    fa.Rows = 1
    fa.FixedRows = 1
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.SelectionMode = flexSelectionListBox
    fa.ExplorerBar = flexExMove ' can't sort when bound
    fa.AutoSearch = flexSearchFromCursor
    fa.Editable = True
    
    ' connect the control to the variant array returned by ADO
    fa.BindToArray arr
    
    ' set up headers and formats
    Dim c&
    For c = 0 To fa.Cols - 1
        fa.Cell(flexcpText, 0, c) = rs.Fields(c).name
        If rs.Fields(c).Type = adDate Then
            fa.ColDataType(c) = flexDTDate
            fa.ColFormat(c) = "Medium Date"
        End If
    Next
    
    ' do an autosize
    fa.AutoSize 0, fa.Cols - 1, , 100
End Sub

Sub DemoCell()
    
    ' reset the control
    SetDefaults fa
    
    ' set the properties we want
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.ExplorerBar = flexExSortShowAndMove
    
    ' fill the control with data
    fa.Rows = 150
    fa.FixedRows = 1
    fa.Cols = 4
    fa.ColFormat(1) = "#,###"
    fa.ColDataType(3) = flexDTDate
    fa.ColFormat(3) = "Medium Date"
    FillColumn fa, 0, slFileNames
    FillColumn fa, 1, slFileSizes
    FillColumn fa, 2, slFileTypes
    FillColumn fa, 3, slFileDates
    
    ' use the cell property to put pictures and
    ' change the background color on column 0:
    ' this is quick: there's no need to select the cells first
    fa.Cell(flexcpPicture, 1, 0, fa.Rows - 1) = imgCell
    fa.Cell(flexcpPictureAlignment, 1, 0, fa.Rows - 1) = flexPicAlignLeftCenter
    fa.Cell(flexcpBackColor, 1, 0, fa.Rows - 1) = RGB(255, 255, 0)

    ' change the font on one cell
    fa.Cell(flexcpFontBold, 1, 0) = True
    fa.Cell(flexcpFontSize, 1, 0) = 7
    fa.Cell(flexcpFontUnderline, 1, 0) = True
    
    ' copy the new font in cell (1,0) to the entire column 0.
    ' this is double-quick: there's no need to select the cells first,
    ' plus you can change the entire font at once.
    ' note the syntax: you need to write "Set fa.Cell..."
    Set fa.Cell(flexcpFont, 1, 0, fa.Rows - 1) = fa.Cell(flexcpFont, 1, 0)
    
    ' make column widths equal
    fa.ColWidth(-1) = fa.ClientWidth / fa.Cols
    
End Sub

Sub DemoDragDrop()
    
    ' reset the control
    SetDefaults fa
    
    ' set the properties we want
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.SelectionMode = flexSelectionListBox
    
    ' fill the control with data
    fa.Rows = 50
    fa.FixedRows = 1
    fa.Cols = 4
    fa.ColFormat(1) = "#,###"
    fa.ColDataType(3) = flexDTDate
    fa.ColFormat(3) = "Medium Date"
    FillColumn fa, 0, slFileNames
    FillColumn fa, 1, slFileSizes
    FillColumn fa, 2, slFileTypes
    FillColumn fa, 3, slFileDates
    
    ' do an autosize
    fa.AutoSize 0, fa.Cols - 1, , 300
    
    ' let user play with it:
    ' - we'll trap the BeforeMouseDown event to allow
    '   multi-row drag-drop
    
End Sub

Sub DemoEditMask()
    
    ' reset the control
    SetDefaults fa
    
    ' set the properties we want
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.Editable = True
    
    ' fill the control with data
    fa.Rows = 50
    fa.FixedRows = 1
    fa.Cols = 6
    
    FillColumn fa, 0, slFirstNames
    FillColumn fa, 1, slLastNames
    
    FillColumn fa, 2, slPhones
    fa.Cell(flexcpText, 0, 2) = "Office Phone"
    fa.ColEditMask(2) = "(###) 000-0000"
    
    FillColumn fa, 3, slPhones
    fa.Cell(flexcpText, 0, 3) = "Home Phone"
    fa.ColEditMask(3) = "(###) 000-0000"
    
    FillColumn fa, 4, slCards
    fa.ColEditMask(4) = "0000 0000 0000 0000;;*"
    
    FillColumn fa, 5, slFileDates
    fa.ColFormat(5) = "Medium Date"
    fa.ColEditMask(5) = "00->LLL-00"
    
    ' do an AutoSize
    fa.AutoSize 0, fa.Cols - 1, , 300
    
    ' let user play with it:

End Sub

Sub DemoExplorerBar()
    
    ' reset the control
    SetDefaults fa
    
    ' set the properties we want
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.SelectionMode = flexSelectionListBox
    
    ' fill the control with data
    fa.Rows = 50
    fa.FixedRows = 1
    fa.Cols = 4
    fa.ColFormat(1) = "#,###"
    fa.ColDataType(3) = flexDTDate
    fa.ColFormat(3) = "Medium Date"
    FillColumn fa, 0, slFileNames
    FillColumn fa, 1, slFileSizes
    FillColumn fa, 2, slFileTypes
    FillColumn fa, 3, slFileDates
    
    ' do an autosize
    fa.AutoSize 0, fa.Cols - 1, , 300
    
    ' let user play with it:
    ' - sort by clicking column headers,
    ' - rearrange columns by dragging their headers.
    ' here's the code:
    fa.ExplorerBar = flexExSortShowAndMove
    
    
End Sub


Sub DemoScrollTips()
    
    ' use explorer bar demo as a basis
    DemoExplorerBar
    
    ' turn on ScrollTips
    fa.ScrollTrack = False
    fa.ScrollTips = True

End Sub

Sub DemoSubtotals()
    
    ' reset the control
    SetDefaults fa
    
    ' set the properties we want
    fa.Rows = 50
    fa.Cols = 4
    fa.FixedRows = 1
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.ExplorerBar = flexExMove ' can't sort when bound
    fa.Editable = True
    
    ' add data
    FillColumn fa, 0, slProduct
    FillColumn fa, 1, slAssociate
    FillColumn fa, 2, slRegion
    FillColumn fa, 3, slSales
    
    ' sort and total
    DemoSubtotalsSort
End Sub

Sub DemoSubtotalsSort()
    
    ' sort
    fa.Select 1, 0, 1, fa.Cols - 1
    fa.Sort = flexSortGenericAscending
    
    ' insert subtotals
    fa.OutlineCol = 0
    fa.OutlineBar = flexOutlineBarComplete
    fa.SubtotalPosition = flexSTAbove
    fa.Subtotal flexSTClear
    fa.Subtotal flexSTSum, -1, 3, "#,", 1, vbWhite, True
    fa.Subtotal flexSTSum, 0, 3, "#,", &H404040, vbWhite, True
    fa.Subtotal flexSTSum, 1, 3, "#,", &H808080, vbWhite, False
    fa.Subtotal flexSTSum, 2, 3, "#,", &HC0C0C0, vbWhite, False
    
    ' auto size
    fa.AutoSize 0, fa.Cols - 1

    ' merge
    fa.MergeCells = flexMergeRestrictAll
    fa.MergeCol(0) = True
    fa.MergeCol(1) = True
    fa.MergeCol(2) = True
    fa.MergeCol(3) = True
End Sub

Sub DemoTooltips()
    
    ' use explorer bar demo as a basis
    DemoExplorerBar
    
    ' set tooltiptext in response to mousemove event
    ' ...
    
End Sub
Sub DemoTree()
    
    ' reset the control
    SetDefaults fa
    fa.Redraw = False
    
    ' set the properties we want
    fa.Rows = 1
    fa.FixedRows = 1
    fa.Cols = 3
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.OutlineBar = flexOutlineBarComplete
    fa.OutlineCol = 0
    fa.SubtotalPosition = flexSTAbove
    
    ' fill the control with data
    fa.Cell(flexcpText, 0, 0) = "Heading"
    fa.Cell(flexcpText, 0, 1) = "Date"
    fa.Cell(flexcpText, 0, 2) = "Time"
    Dim i%, j%
    While fa.Rows < 150
        
        ' decide randomly whether to add a subtotal
        If fa.Rows <= 2 Or Rnd() < 0.4 Then
        
            ' add an item, make it a subtotal
            fa.AddItem "Branch Level " & i
            fa.IsSubtotal(fa.Rows - 1) = True
            fa.RowOutlineLevel(fa.Rows - 1) = i
            fa.Cell(flexcpPicture, fa.Rows - 1, 0) = imgFolder
            j = i
            
            ' decide whether to go deeper or shallower
            If Rnd() < 0.5 And i < 10 Then
                i = i + 1
            ElseIf Rnd() < 0.5 And i > 0 Then
                i = i - 1
            End If
        
        ' add a regular item
        Else
            fa.AddItem "Data on Level " & j & vbTab & Date & vbTab & Time
            fa.Cell(flexcpPicture, fa.Rows - 1, 0) = imgItem
        End If
    Wend
    
    ' add some more data just for kicks
    For i = 1 To 10
        fa.AddItem "Data on Level " & j & vbTab & Date & vbTab & Time
        fa.Cell(flexcpPicture, fa.Rows - 1, 0) = imgItem
    Next
    
    ' do an autosize
    fa.AutoSize 0, 1, , 300
    fa.Redraw = True
    
End Sub

Sub DemoDDAuto()
    
    ' reset the control
    SetDefaults fa
    
    ' fill the control with data
    fa.Rows = 50
    fa.FixedRows = 1
    fa.Cols = 4
    fa.ColFormat(1) = "#,###"
    fa.ColDataType(3) = flexDTDate
    fa.ColFormat(3) = "Medium Date"
    FillColumn fa, 0, slFileNames
    FillColumn fa, 1, slFileSizes
    FillColumn fa, 2, slFileTypes
    FillColumn fa, 3, slFileDates
    
    ' set the properties we want
    fa.OLEDragMode = flexOLEDragAutomatic
    fa.OLEDropMode = flexOLEDropAutomatic
    
End Sub


Sub DemoDDManual()
    
    ' reset the control
    SetDefaults fa
    
    ' fill the control with data
    fa.Rows = 50
    fa.FixedRows = 1
    fa.Cols = 4
    fa.ColFormat(1) = "#,###"
    fa.ColDataType(3) = flexDTDate
    fa.ColFormat(3) = "Medium Date"
    FillColumn fa, 0, slFileNames
    FillColumn fa, 1, slFileSizes
    FillColumn fa, 2, slFileTypes
    FillColumn fa, 3, slFileDates
    
    ' set the properties we want
    fa.OLEDragMode = flexOLEDragManual
    fa.OLEDropMode = flexOLEDropManual
    
End Sub



Sub FillColumn(fa As VSFlexGrid, ByVal c As Long, ByVal s As String)
    Dim r&, i&, cnt&
    ReDim lst(0) As String
    
    ' build list
    cnt = 0
    i = InStr(s, "|")
    While i > 0
        lst(cnt) = Left(s, i - 1)
        s = Mid(s, i + 1)
        cnt = cnt + 1
        ReDim Preserve lst(cnt) As String
        i = InStr(s, "|")
    Wend
    lst(cnt) = s
    
    ' set values
    fa.Cell(flexcpText, 0, c) = lst(0)
    For r = fa.FixedRows To fa.Rows - 1
        i = (Rnd() * 1000) Mod cnt + 1
        fa.Cell(flexcpText, r, c) = lst(i)
    Next
    
    ' do an autosize
    fa.AutoSize c, , , 300
End Sub
Sub StringReplace(s$, ByVal f$, ByVal r$)
    Dim i&, l&

    ' sanity check
    l = Len(f)
    If l = 0 Then Exit Sub

    ' find and replace
    i = 0
    Do
        i = InStr(i + 1, s, f)
        If i = 0 Then Exit Do
        s = Left(s, i - 1) & r & Mid(s, i + l)
    Loop
End Sub

Sub SetDefaults(fa As VSFlexGrid)
    
    With fa
        .BindToArray Null
        .Rows = 0
        .Cols = 0
        .ScrollTrack = False
        .ExplorerBar = flexExNone
        .AutoSearch = flexSearchNone
        .Editable = False
        .AllowUserResizing = flexResizeNone
        .SelectionMode = flexSelectionFree
        .OutlineBar = flexOutlineBarNone
        .OLEDragMode = flexOLEDragManual
        .OLEDropMode = flexOLEDropNone
        .ScrollTips = False
        .ToolTipText = ""
    End With
    
End Sub

Private Sub btnGo_Click()
    Dim i%
    
    ' no demo for headings
    If lstMenu.IsSubtotal(lstMenu.Row) Then Exit Sub
    
    ' remember which demo we're showing
    i = lstMenu.RowData(lstMenu.Row)
    CurrentDemo = Demo(i).name
    
    ' update form caption
    Caption = "VideoSoft VSFLEX6 Demo - " & CurrentDemo
    fa.SetFocus
    
    ' go show it
    If InStr(CurrentDemo, "Mask") > 0 Then
        DemoEditMask
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "Popup") > 0 Then
        DemoPopup
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "Combo") > 0 Then
        DemoCombo
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "Dates") > 0 Then
        DemoFormat
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "Cell") > 0 Then
        DemoCell
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "Drag") > 0 Then
        DemoDragDrop
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "Search") > 0 Then
        DemoAutoSearch
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "Explorer") > 0 Then
        DemoExplorerBar
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "2D-ADO") > 0 Then
        DemoBindToArrayADO
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "3D-VB") > 0 Then
        DemoBind3D
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "2D-VB") > 0 Then
        DemoBind2D
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "Synch") > 0 Then
        frmSynch.lblHelp = Demo(i).help
        frmSynch.Show 1
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "another") > 0 Then
        frmSynFa.lblHelp = Demo(i).help
        frmSynFa.Show 1
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "Sales Report") > 0 Then
        DemoSubtotals
        Exit Sub
    End If

    If InStr(CurrentDemo, "Tree") > 0 Then
        DemoTree
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "Automatic") > 0 Then
        DemoDDAuto
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "Manual") > 0 Then
        DemoDDManual
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "Tooltips") > 0 Then
        DemoTooltips
        Exit Sub
    End If
    
    If InStr(CurrentDemo, "ScrollTips") > 0 Then
        DemoScrollTips
        Exit Sub
    End If
    
    ' oops...
    MsgBox "Sorry, '" & CurrentDemo & "' has not been implemented yet."

End Sub

Sub DemoBind3D()
    Static pg%
    
    ' dimension an array for us to bind to
    ' notice that is is declared as static so it won't go out of scope and die
    Static arr(11, 3, 50)
    
    ' fill out the array with some data (first time only, it is static)
    Dim i&, j&, k&
    If IsEmpty(arr(0, 0, 0)) Then
        For i = 0 To 11
            For j = 0 To 3
                For k = 0 To 50
                    arr(i, j, k) = "Page " & i & ", Cell(" & j & "," & k & ")"
                Next
            Next
        Next
    End If
    
    ' reset the control
    SetDefaults fa
    
    ' set the properties we want
    fa.Rows = 1
    fa.FixedRows = 1
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.ExplorerBar = flexExMove ' can't sort when bound
    fa.Editable = True
    
    ' bind to the array
    ' the parameters are used to select which dimensions correspond to rows, columns, and pages
    ' in this case, rows are in dimension 2 (51 of them)
    '               cols are in dimension 1 ( 4 of them)
    '              pages are in dimension 0 (12 of them)
    fa.BindToArray arr, 2, 1, 0, pg
    
    ' select next page for next time
    pg = pg + 1
    If pg > 11 Then pg = 0

    ' set up headers and formats
    For i = 0 To fa.Cols - 1
        fa.Cell(flexcpText, 0, i) = "Col " & i
    Next
    
    ' make column widths equal
    fa.ColWidth(-1) = fa.ClientWidth / fa.Cols
End Sub

Sub DemoBind2D()
    
    ' dimension an array for us to bind to
    ' notice that is is declared as static so it won't go out of scope and die
    Static arr(3, 50)
    
    ' fill out the array with some data (first time only, it is static)
    Dim i&, j&
    If IsEmpty(arr(0, 0)) Then
        For i = 0 To 3
            For j = 0 To 50
                arr(i, j) = "arr(" & i & "," & j & ")"
            Next
        Next
    End If
    
    ' reset the control
    SetDefaults fa
    
    ' set the properties we want
    fa.Rows = 1
    fa.FixedRows = 1
    fa.AllowUserResizing = flexResizeBoth
    fa.ExtendLastCol = True
    fa.ExplorerBar = flexExMove ' can't sort when bound
    fa.Editable = True
    
    ' bind to the array
    fa.BindToArray arr

    ' set up headers and formats
    For i = 0 To fa.Cols - 1
        fa.Cell(flexcpText, 0, i) = "Col " & i
    Next
    
    ' make column widths equal
    fa.ColWidth(-1) = fa.ClientWidth / fa.Cols
End Sub



Private Sub Command1_Click()
    SaveFlexGridToHTML fa, "c:\temp\fa.htm"
End Sub

Private Sub fa_AfterMoveColumn(ByVal Col As Long, Position As Long)

    ' update subtotals
    If InStr(CurrentDemo, "Sales Report") > 0 Then
        DemoSubtotalsSort
    End If

End Sub

Private Sub fa_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
    
    ' trap this to implement drag/drop
    If InStr(CurrentDemo, "Drag") > 0 Then
        
        ' use shift to drag (ctrl selects)
        If Shift = 1 Then
            
            ' build a list of what we're dragging
            Dim i&
            fa.Tag = ""
            'For i = fa.FixedRows To fa.Rows - 1
            '    If fa.IsSelected(i) Then fa.Tag = fa.Tag & vbCrLf & vbTab & fa.Cell(flexcpText, i, 0)
            'Next
            For i = 0 To fa.SelectedRows - 1
                fa.Tag = fa.Tag & vbCrLf & vbTab & fa.Cell(flexcpText, fa.SelectedRow(i), 0)
            Next
            
            ' start dragging
            fa.Drag
            
            ' cancel remaining mouse events
            Cancel = True
        End If
    End If

End Sub

Private Sub fa_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    
    ' don't move sales figures
    If InStr(CurrentDemo, "Sales Report") > 0 Then
        If Col = fa.Cols - 1 Then Position = -1
    End If

End Sub

Private Sub fa_BeforeScrollTip(ByVal Row As Long)
    
    fa.ScrollTipText = " Row " & Row & ": " & fa.Cell(flexcpTextDisplay, Row, 0) & " "

End Sub

Private Sub fa_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim pt As POINTAPI
    
    ' get popup window position
    pt.X = fa.ColPos(Col) / Screen.TwipsPerPixelX
    pt.Y = (fa.RowPos(Row) + fa.RowHeight(Row)) / Screen.TwipsPerPixelY
    ClientToScreen fa.hwnd, pt
    
    ' show date popup
    If fa.ColDataType(Col) = flexDTDate Then
        With frmDate
            .lblRow = Row
            .lblCol = Col
            .Tag = fa.Cell(flexcpText, Row, Col)
            .Move pt.X * Screen.TwipsPerPixelX, pt.Y * Screen.TwipsPerPixelY
            .Show
        End With
        Exit Sub
    End If
    
    ' show file popup
    If InStr(fa.Cell(flexcpText, 0, Col), "File") > 0 Then
        With frmFile
            .lblRow = Row
            .lblCol = Col
            .Tag = fa.Cell(flexcpText, Row, Col)
            .Move pt.X * Screen.TwipsPerPixelX, pt.Y * Screen.TwipsPerPixelY
            .Show
        End With
        Exit Sub
    End If

End Sub

Private Sub fa_ChangeEdit()
    Debug.Print "change edit to: ["; fa.EditText; "]"
End Sub

Private Sub fa_DragDrop(Source As Control, X As Single, Y As Single)
    
    MsgBox "You dragged the following records: " & fa.Tag

End Sub








Private Sub fa_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Select Case State
        Case vbEnter
            fa.DragIcon = imgDrag
        Case vbLeave
            fa.DragIcon = imgNope
    End Select
End Sub




Private Sub fa_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Debug.Print Button; Shift; X; Y
    
    ' show regular tooltip
    If InStr(CurrentDemo, "Tooltip") > 0 Then
        Dim r&, c&
        r = fa.MouseRow
        c = fa.MouseCol
        If r >= 0 And c >= 0 Then
            fa.ToolTipText = "This is " & fa.Cell(flexcpText, r, c)
        Else
            fa.ToolTipText = ""
        End If
    End If

End Sub


Private Sub fa_OLEDragDrop(Data As VSFlex7lCtl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    ' if we're dropping files, open the file and paste contents
    If InStr(CurrentDemo, "Manual") > 0 Then
        If Data.FileCount > 0 Then
            On Error Resume Next
            Open Data.Files(0) For Input As #1
            Dim i%, r%, c%, s$
            r = fa.MouseRow
            c = fa.MouseCol
            For i = 0 To 10
                Line Input #1, s
                fa.Cell(flexcpText, r + i, c) = s
            Next
            Close #1
        Else
            MsgBox "For this demo, we only accept dropping files..."
        End If
    End If
End Sub




Private Sub fa_OLEDragOver(Data As VSFlex7lCtl.VSDataObject, Effect As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, State As Integer)

    ' if we're dropping files, open the file and paste contents
    If InStr(CurrentDemo, "Manual") > 0 Then
        Effect = vbDropEffectNone
        If Data.FileCount > 0 Then Effect = vbDropEffectCopy
    End If

End Sub


Private Sub Form_Load()
    Dim i%, j%
    
    ' fill demo structure
    i = 0
    ReDim Preserve Demo(i)
    Demo(i).name = "Editing: Masks"
    Demo(i).help = "This demo shows how you can use input masks to control text input. " & vbCrLf & _
                   "The code sets the ColEditMask property when it populates the grid (it does not use the BeforeEdit event)."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Editing: Popups"
    Demo(i).help = "This demo shows how you can create pop-up edit windows. " & vbCrLf & _
                   "The code sets the ColComboList property to an ellipsis ('...') when it populates the grid (it does not use the BeforeEdit event). " & _
                   "Then it traps the CellButtonClick event to show a custom popup editor."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Editing: Combos"
    Demo(i).help = "This demo shows how you can create multi-column drop-downs and combos. " & vbCrLf & _
                   "The code sets the ColComboList property when it populates the grid (it does not use the BeforeEdit event)."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Formatting: Numbers, Dates, Bools"
    Demo(i).help = "This demo shows how you can format numbers, dates, and boolean values. " & vbCrLf & _
                   "The code sets the ColDataType and ColFormat properties."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Formatting: Cell Property"
    Demo(i).help = "This demo shows how you can use the new Cell property to set fonts and pictures in individual cells. " & vbCrLf & _
                   "The Cell property allows you to get or set any cell property for a cell or range without selecting it first."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Binding to Arrays: 2D-VB Arrays"
    Demo(i).help = "This demo shows how you can bind the control to a 2D-VB array using the BindToArray method. " & vbCrLf & _
                   "The control then reads and writes data directly from and to the array."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Binding to Arrays: 3D-VB Arrays"
    Demo(i).help = "This demo shows how you can bind the control to a 3D-VB array using the BindToArray method. " & vbCrLf & _
                   "The control selects a random 'page' to display each time you run this demo. You may edit the entries."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Binding to Arrays: 2D-ADO Arrays"
    Demo(i).help = "This demo shows how you can bind the control to an ADO recordset using the BindToArray method. " & vbCrLf & _
                   "You must have a DataSource named 'Northwind' for this to work."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Binding to Arrays: Synchronized controls"
    Demo(i).help = "This demo shows four FlexGrid controls that are synchronized." & vbCrLf & _
                   "Data is shared through a common array binding." & vbCrLf & _
                   "Scrolling and selecting are synchronized with some VB code." & vbCrLf & _
                   "Edit one control and watch the others change."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Binding to Arrays: Bind to another FlexGrid"
    Demo(i).help = "This demo shows four FlexGrid controls that are synchronized." & vbCrLf & _
                   "Type values on the main control. The others will display the totals." & vbCrLf & _
                   "Scrolling is synchronized with some VB code."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Goodies: AutoSearch"
    Demo(i).help = "This demo shows how the new AutoSearch property works. " & vbCrLf & _
                   "Just start typing and the control will look for the data dynamically."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Goodies: Explorer Bar"
    Demo(i).help = "This demo shows how the new ExplorerBar property works. " & vbCrLf & _
                   "Click on any column header to sort the column in ascending or descending order. Drag column headers around to reorder them."

    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Goodies: Drag-Drop Multiple Rows"
    Demo(i).help = "This demo shows how to implement drag-and-drop for multiple rows. " & vbCrLf & _
                   "Select rows by ctrl-clicking them, then shift-drag them. See the BeforeMouseDown event to find out how it's done."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Outlines/Totals: Sales Report"
    Demo(i).help = "This demo shows how to implement an informative sales summary." & vbCrLf & _
                   "Drag and drop columns to get different views of the sales report."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Outlines/Totals: Custom Tree"
    Demo(i).help = "This demo shows how to build a custom outline tree without using subtotals." & vbCrLf & _
                   "The tree is similar to the one used to create the main menu on this demo."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Ole Drag Drop: Automatic OleDD"
    Demo(i).help = "This demo shows how to implement Ole drag and drop automatically." & vbCrLf & _
                   "Drag data to other programs or to other parts of the grid by dragging the edge of the selection."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Ole Drag Drop: Manual OleDD"
    Demo(i).help = "This demo shows how to implement manual Ole drag and drop." & vbCrLf & _
                   "Drag text files from the windows Explorer and drop them on the control. It will display the ten first lines of the file."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Tooltips: Standard Tooltips"
    Demo(i).help = "This demo shows how to implement ToolTips using VB's 'ToolTipText' property." & vbCrLf & _
                   "Place the mouse over any cell and wait a couple of seconds."
    
    i = i + 1
    ReDim Preserve Demo(i)
    Demo(i).name = "Tooltips: VideoSoft ScrollTips"
    Demo(i).help = "This demo shows how to implement 'ScrollTips'." & vbCrLf & _
                   "Use the scrollbars to scroll the control vertically and you will see s ToolTip on the scrollbar, telling you where you are."
    
    ' build menu
    Dim s$, slast$
    With lstMenu
    .OutlineBar = flexOutlineBarSimple
    .SubtotalPosition = flexSTAbove
    For i = 0 To UBound(Demo)
        
        ' quit when done
        If Demo(i).name = "" Then Exit For
        
        ' add item
        j = InStr(Demo(i).name, ":")
        If j > 0 Then
            s = Trim(Left(Demo(i).name, j))
            If s <> slast Then
                slast = s
                .AddItem s
                .IsSubtotal(.Rows - 1) = True
                .Cell(flexcpFontBold, .Rows - 1, 0) = True
                .Cell(flexcpFontUnderline, .Rows - 1, 0) = True
                .Cell(flexcpPicture, .Rows - 1, 0) = imgFolder
            End If
            Demo(i).name = Trim(Mid(Demo(i).name, j + 1))
            .AddItem Demo(i).name
            .Cell(flexcpPicture, .Rows - 1, 0) = imgItem
            .Cell(flexcpForeColor, .Rows - 1, 0) = lblHelp.ForeColor
            .RowData(.Rows - 1) = i
        End If
    Next
    End With
    
    ' start
    lstMenu.Row = 1
    
End Sub


Private Sub Form_Resize()
    Dim X!, edge!
    
    ' set label width
    edge = lblHelp.Top
    X = ScaleWidth - lblHelp.Left - edge
    If X > 0 Then lblHelp.Width = X
    
    ' set list and fa height
    X = ScaleHeight - lstMenu.Top - edge
    If X > 0 Then lstMenu.Height = X: fa.Height = X

    ' set fa width
    X = ScaleWidth - fa.Left - edge
    If X > 0 Then fa.Width = X
    
End Sub




Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


Private Sub lblHelp_GotFocus()
    fa.SetFocus
End Sub




Private Sub lblHelp_KeyDown(KeyCode As Integer, Shift As Integer)
    
    KeyCode = 0

End Sub


Private Sub lblHelp_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub lstMenu_DblClick()
    If lstMenu.IsSubtotal(lstMenu.Row) Then
        lstMenu.Outline 0
        lstMenu.IsCollapsed(lstMenu.Row) = flexOutlineExpanded
    Else
        btnGo_Click
    End If
End Sub

Private Sub lstMenu_RowColChange()
    Dim i%
    
    ' no help for headings
    If lstMenu.IsSubtotal(lstMenu.Row) Then Exit Sub
    
    ' show help for this item
    i = lstMenu.RowData(lstMenu.Row)
    lblHelp = Demo(i).help
    
    ' fix button caption
    btnGo.Caption = "Show Me" & vbCrLf & Demo(i).name

End Sub

