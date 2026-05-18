VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Begin VB.Form frmPreInvoice 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pre Invoice"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   12540
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3300
      Top             =   8640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8595
      Left            =   2640
      ScaleHeight     =   8565
      ScaleWidth      =   9225
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   9255
      Begin VSFlex7LCtl.VSFlexGrid flexTitle 
         Height          =   675
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   7455
         _cx             =   13150
         _cy             =   1191
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   16777215
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   0
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
         FormatString    =   $"frmPreInvioce.frx":0000
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
         ShowComboButton =   1
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
         Height          =   5865
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   8175
         _cx             =   14420
         _cy             =   10345
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
         BackColorFixed  =   16777215
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   -2147483641
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   0
         GridColorFixed  =   0
         TreeColor       =   -2147483639
         FloodColor      =   0
         SheetBorder     =   0
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
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
         FormatString    =   $"frmPreInvioce.frx":00D6
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         ShowComboButton =   1
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
      Begin VSFlex7LCtl.VSFlexGrid flexCust 
         Height          =   495
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   7455
         _cx             =   13150
         _cy             =   873
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   16777215
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   16777215
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   0
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
         FormatString    =   $"frmPreInvioce.frx":015F
         ScrollTrack     =   0   'False
         ScrollBars      =   0
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
         ShowComboButton =   1
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
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9120
      Left            =   0
      ScaleHeight     =   9090
      ScaleWidth      =   2625
      TabIndex        =   11
      Top             =   0
      Width           =   2655
      Begin VB.ComboBox cboCompany 
         BackColor       =   &H00FFC0C0&
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Option2 
         Caption         =   "แสดงใบคุม"
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Top             =   7440
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "ไม่ต้องแสดงใบคุม"
         Height          =   315
         Left            =   240
         TabIndex        =   25
         Top             =   7680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtDocNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "Export to Excel"
         Height          =   435
         Left            =   180
         TabIndex        =   8
         Top             =   5940
         Width           =   2295
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   1320
         Width           =   2295
      End
      Begin VB.ComboBox cboJobStatus 
         Height          =   315
         ItemData        =   "frmPreInvioce.frx":0235
         Left            =   180
         List            =   "frmPreInvioce.frx":0237
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2280
         Width           =   2325
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   435
         Left            =   180
         TabIndex        =   7
         Top             =   5460
         Width           =   2295
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   435
         Left            =   180
         TabIndex        =   10
         Top             =   6900
         Width           =   2295
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Print To Printer"
         Height          =   435
         Left            =   180
         TabIndex        =   9
         Top             =   6420
         Width           =   2295
      End
      Begin VB.TextBox txtJobNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   4920
         Width           =   2295
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Text            =   "*"
         Top             =   4320
         Width           =   2295
      End
      Begin DateCtl.Date dtFrom 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   2880
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin DateCtl.Date dtTo 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   3600
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท :"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   120
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบเมนิฟเส"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   24
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   20
         Top             =   1080
         Width           =   660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ปฏิบัติงานตั้งแต่"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   19
         Top             =   2640
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   18
         Top             =   3240
         Width           =   495
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   180
         Top             =   8190
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   180
         Top             =   8490
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   180
         Top             =   8790
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   180
         Top             =   9090
         Width           =   195
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "งานใหม่"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   450
         TabIndex        =   17
         Top             =   8220
         Width           =   570
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พิมพ์ Manifest แล้ว"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   450
         TabIndex        =   16
         Top             =   8520
         Width           =   1335
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปิดงานแล้ว"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   450
         TabIndex        =   15
         Top             =   8820
         Width           =   780
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   450
         TabIndex        =   14
         Top             =   9120
         Width           =   480
      End
      Begin VB.Line Line1 
         X1              =   180
         X2              =   2490
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบสั่งงาน,เลขที่ใบคุม :"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   4680
         Width           =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   12
         Top             =   4080
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmPreInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer
Dim SaleStaff_ As BWGCompanyStaff

Dim mDocTitle$
Dim mDocTitleCust$
Dim mBackColor&

Const mForeColor& = vbWhite
Const mFontName$ = "MS Sans Serif"
Const mFontSize = "8"
Const mCellAlignment = flexAlignCenterCenter

''''' *************** Print Picture Box API *********************
Private Declare Function SendMessage Lib "user32" Alias _
         "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
         ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317
Private Const PRF_CLIENT = &H4&    ' Draw the window's client area
Private Const PRF_CHILDREN = &H10& ' Draw all visible child windows
Private Const PRF_OWNED = &H20&    ' Draw all owned windows
''''' *************** End Print Picture Box *********************

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Sub SaveFormImageToFile(ByRef ContainerForm As Form, ByRef PictureBoxControl As PictureBox, ByVal ImageFileName As String)
  Dim FormInsideWidth As Long
  Dim FormInsideHeight As Long
  Dim PictureBoxLeft As Long
  Dim PictureBoxTop As Long
  Dim PictureBoxWidth As Long
  Dim PictureBoxHeight As Long
  Dim FormAutoRedrawValue As Boolean
  
  With PictureBoxControl
    'Set PictureBox properties
    .Visible = False
    .AutoRedraw = True
    .Appearance = 0 ' Flat
    .AutoSize = False
    .BorderStyle = 0 'No border
    
    'Store PictureBox Original Size and location Values
    PictureBoxHeight = .Height: PictureBoxWidth = .Width: PictureBoxLeft = .Left: PictureBoxTop = .Top
    
    'Make PictureBox to size to inside of form.
    .Align = vbAlignTop: .Align = vbAlignLeft
    DoEvents
    
    FormInsideHeight = .Height: FormInsideWidth = .Width
    
    'Restore PictureBox Original Size and location Values
    .Align = vbAlignNone
    .Height = FormInsideHeight: .Width = FormInsideWidth: .Left = PictureBoxLeft: .Top = PictureBoxTop
    
    FormAutoRedrawValue = ContainerForm.AutoRedraw
    ContainerForm.AutoRedraw = False
    DoEvents
    
    'Copy Form Image to Picture Box
    BitBlt .hdc, 0, 0, FormInsideWidth / Screen.TwipsPerPixelX, FormInsideHeight / Screen.TwipsPerPixelY, ContainerForm.hdc, 0, 0, vbSrcCopy
    DoEvents
    SavePicture .Image, ImageFileName
    DoEvents
    
    ContainerForm.AutoRedraw = FormAutoRedrawValue
    DoEvents
  End With
End Sub

Private Sub FlexGrid_To_Excel(ByRef Flexgrid As VSFlexGrid, ByVal Rows As Integer, ByVal FCols As Integer, ByVal LCols As Integer, _
    Optional ByVal GridStyle As Integer = 0, Optional ByVal WorkSheetName As String = "")
    
    Dim objXL As Object     '    Dim objXL As New Excel.Application
    Dim wbXL As Object      '    Dim wbXL As New Excel.Workbook
    Dim wsXL As Object      '    Dim wsXL As New Excel.Worksheet
    Dim intRow As Integer ' counter
    Dim intCol As Integer ' counter
    
'    If Not IsObject(objXL) Then
'        MsgBox "You need Microsoft Excel to use this function", vbExclamation, "Print to Excel"
'        Exit Sub
'    End If
    
    'On Error Resume Next is necessary because
    'someone may pass more rows
    'or columns than the flexgrid has
    
    'you can instead check for this,
    'or rewrite the function so that
    'it exports all non-fixed cells
    'to Excel
    
'    On Error Resume Next
    
    ' open Excel
    Set objXL = CreateObject("Excel.Application")
    Set wbXL = objXL.Workbooks.Add
    Set wsXL = objXL.ActiveSheet
    
    ' name the worksheet
    With wsXL
        If Not WorkSheetName = "" Then
            .Name = WorkSheetName
        End If
    End With
        
    ' fill worksheet
    If Not curCustomer_ Is Nothing Then wsXL.Cells(1, 2).Value = curCustomer_.CustomerCode & " " & curCustomer_.CustomerName
    For intRow = 2 To Rows + 1
        For intCol = FCols To LCols
            With Flexgrid
                wsXL.Cells(intRow, intCol).Value = .TextMatrix(intRow - 1, intCol - 1) & " "
            End With
        Next
    Next
    
    ' format the look
    For intCol = FCols To LCols
        wsXL.Columns(intCol).AutoFit
        'wsXL.Columns(intCol).AutoFormat (1)
        wsXL.Range("a1", Right(wsXL.Columns(LCols).AddressLocal, 1) & Rows + 1).AutoFormat 0
    Next
    
    objXL.Visible = True
    Set wsXL = Nothing
    Set wbXL = Nothing
    Set objXL = Nothing
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    On Error GoTo ErrD
    
'''********************* Original export to Excel file ********************************
'    Dim AppExcel As Object
'    Dim ExcelFile As String
'    Set AppExcel = CreateObject("Excel.Application")
'    ExcelFile = App.Path & "\" & "PreInv" & IIf(txtCustomer.Text <> "", "_" & Replace(Replace(txtCustomer.Text, "/", "-"), "*", ""), "") & "_" & Replace(dtFrom.ValueYMD & "_" & dtTo.ValueYMD, "/", "-") & ".xls"

'    flexItems.FixedCols = 1
'    flexItems.FixedRows = 1
'
'    flexItems.FixedCols = flexItems.FixedCols
'    flexItems.FixedRows = 1
'
'    flexItems.SaveGrid ExcelFile, flexFileTabText
'    AppExcel.Workbooks.open ExcelFile
'
'    AppExcel.Visible = True
'    Set AppExcel = Nothing
'
'    flexItems.FixedCols = 0
'    flexItems.FixedRows = 0
'''********************* End Original export to Excel file ********************************

''''****************** Flexgrid copy to Clipboard and past to Excel Object ************************
'    Dim xlObject As Excel.Application
'    Dim xlWB As Excel.Workbook
'
'    Set xlObject = CreateObject("Excel.Application")
'
'    'This Adds a new woorkbook, you could open the workbook from file also
'    Set xlWB = xlObject.Workbooks.Add
'
'    Clipboard.Clear 'Clear the Clipboard
'    With flexItems
'        'Select Full Contents (You could also select partial content)
'        .Col = 1               'From first column
'        .Row = 0               'From first Row (header)
'        .ColSel = .Cols - 2    'Select all columns
'        .RowSel = .Rows - 1    'Select all rows
'        Clipboard.SetText .Clip 'Send to Clipboard
'    End With
'
'    With xlObject.ActiveWorkbook.ActiveSheet
'        .Range("A1").Select 'Select Cell A1 (will paste from here, to different cells)
'        .Paste              'Paste clipboard contents
'    End With
'
'    ' This makes Excel visible
'    xlObject.Visible = True
    
''''****************** End Flexgrid copy to Clipboard and past to Excel Object ************************
    
    FlexGrid_To_Excel flexItems, flexItems.Rows - 1, 2, flexItems.Cols - 1, 0, "PreInv"
    
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub cmdPreview_Click()
Dim AppExcel As Object
Dim ExcelFile As String
'    Set AppExcel = CreateObject("Excel.Application")
'    ExcelFile = App.Path & "\" & "PreInv" & IIf(txtCustomer.Text <> "", "_" & Replace(Replace(txtCustomer.Text, "/", "-"), "*", ""), "") & "_" & Replace(dtFrom.ValueYMD & "_" & dtTo.ValueYMD, "/", "-") & ".xls"
'
'    flexItems.FixedCols = 1
'    flexItems.FixedRows = 1
'
'    flexItems.FixedCols = flexItems.FixedCols
'    flexItems.FixedRows = 1
'
'    flexItems.SaveGrid ExcelFile, flexFileTabText
'    AppExcel.Workbooks.open ExcelFile
'
'    AppExcel.Visible = True
'    Set AppExcel = Nothing
'
'    flexItems.FixedCols = 0
'    flexItems.FixedRows = 0

'       Dim rv As Long
'       Picture2.SetFocus  ' So that the button doesn't look pressed
'       Picture2.AutoRedraw = True
'       rv = SendMessage(Picture2.hwnd, WM_PAINT, Picture2.hdc, 0)
'       rv = SendMessage(Picture2.hwnd, WM_PRINT, Picture2.hdc, _
'          PRF_CHILDREN + PRF_CLIENT + PRF_OWNED)
'       Picture2.Picture = Picture2.Image
'       Picture2.AutoRedraw = False
'       Printer.Orientation = vbPRORLandscape   ' 2
'       Printer.Print ""
'       Printer.PaintPicture Picture2.Picture, 0, 0
'       Printer.CurrentY = Picture2.Height + 200   ' get below image
'       Printer.Print "Text on the same page as the image!"
'       Printer.EndDoc

    On Error Resume Next
    Dim OSName As String
    Dim OSBuildNumber As String
    Dim OSVersion As Double
    Dim OSCSDVersion As String
    Dim hKey As Long, NPage As Long, TotalPages As Double
    Dim ModNumber As Integer
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
        REG_PRODUCT_KEY, _
        0&, KEY_QUERY_VALUE, hKey) <> ERROR_SUCCESS _
    Then
        MsgBox "Error opening key."
        Exit Sub
    End If

    ' Get the subkeys' values.
    OSName = GetRegistryValue(hKey, "ProductName")
    OSBuildNumber = GetRegistryValue(hKey, "CurrentBuildNumber")
    OSVersion = CDbl(GetRegistryValue(hKey, "CurrentVersion"))
    OSCSDVersion = GetRegistryValue(hKey, "CSDVersion")

    ' Close the key.
    If RegCloseKey(hKey) <> ERROR_SUCCESS Then
        MsgBox "Error closing key."
    End If
    Dim PrinterDialog As CommonDialog
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlPDPrintSetup
    CommonDialog1.PrinterDefault = False
    CommonDialog1.ShowPrinter
    CommonDialog1.FileName = "PreInv"
    If OSVersion <= 5.1 Then        '*** For Windows XP
        TotalPages = (flexItems.Rows - 1) / 57
        ModNumber = TotalPages
        TotalPages = (1 - (TotalPages - ModNumber)) + (TotalPages - ModNumber) + ModNumber
        For NPage = 1 To TotalPages
            If NPage = 1 Then
                Printer.PaintPicture flexTitle.Picture, 700, 100
                Printer.PaintPicture flexCust.Picture, 700, 900
                Printer.PaintPicture flexItems.Picture, 700, 1500, , , , , , 13800
            ElseIf NPage > 1 Then
                Printer.NewPage
                Printer.PaintPicture flexTitle.Picture, 700, 100
                Printer.PaintPicture flexCust.Picture, 700, 900
                Printer.PaintPicture flexItems.Picture, 700, 1110, , , , , , 200
                Printer.PaintPicture flexItems.Picture, 700, 1810, , , , ((NPage - 1) * 13800) + 150, , 13800
            End If
        Next NPage
        Printer.EndDoc
    Else        '*** Work with Windows 7
        Printer.PaintPicture flexTitle.Picture, 0.5, 0.25
        Printer.PaintPicture flexCust.Picture, 0.5, 0.5
        Printer.PaintPicture flexItems.Picture, 0.5, 0.77
        Printer.EndDoc
        
    End If
    
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrD
    If txtCustomer.Text = "" Or txtCustomer.Text = "*" Then
       Exit Sub
    End If
    flexTitle.Visible = True
    flexCust.Visible = True
    flexItems.Visible = True
    Dim rs As New ADODB.Recordset
    Dim tmpList As ListItem
    Dim x As Long, y As Long
    Dim sqlStr As String, strCond As String, StatusDesc As String
    Dim tmpArr, curColor, subStr As String
    Dim curCondition_ As String
    Dim CarRegisID1 As String, CarRegisID2 As String, LastTimeTableNo As String
    
    Screen.MousePointer = 11
    sqlStr = "SELECT MenifestID,DocNo,MenifestNo,CustomerName,SubTypeDesc,TruckTypeDesc,TransportName1,TransportName2,WorkDate,CloseDate, CloseTime,isClosed,isCanceled,isPrinted,DriverName,JobNo,TimeTableNo,isTrailer " & vbCrLf
    sqlStr = sqlStr & ", CloseStaff = (SELECT TOP(1) StaffFName + ' ' + StaffLName FROM tbCompanyStaff WHERE CloseStaffID = StaffID)" & vbCrLf
    sqlStr = sqlStr & "From vw_MenifestQue "
    If Trim(txtJobNo.Text) <> "" Then
        strCond = " WHERE JobNo Like '%" & Trim(txtJobNo.Text) & "%' Or TimeTableNo Like '%" & Trim(txtJobNo.Text) & "%'"
'    ElseIf Trim(txtDocNo.Text) <> "" Then

    Else
        strCond = " WHERE (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "') " & vbCrLf
        
        If Trim(txtDocNo.Text) <> "" And Replace(Trim(txtDocNo.Text), "*", "") <> "" Then
            strCond = "AND (MenifestNo Like '%" & Replace(Trim(txtDocNo.Text), "*", "%") & "%' OR DocNo Like '%" & Replace(Trim(txtDocNo.Text), "*", "%") & "%' )"
        End If
'        If Trim(txtKey2.Text) <> "" And Replace(Trim(txtKey2.Text), "*", "") <> "" Then
'            strCond = strCond & " And (TransportName1 Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or TransportName2 Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or DriverName Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or CarRegisID Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%')"
'        End If
        If Not curCustomer_ Is Nothing Then
            strCond = strCond & vbCrLf & "AND CustomerID = '" & curCustomer_.ID & "'" & vbCrLf
        End If
        
        If Trim(txtSale.Text) = "" Or Trim(txtSale.Text) = "*" Then Set SaleStaff_ = Nothing
        If Not SaleStaff_ Is Nothing Then
            strCond = strCond & vbCrLf & "AND UnderSaleStaffID = '" & SaleStaff_.ID & "'" & vbCrLf
        End If
        
        Select Case cboJobStatus.ListIndex
            Case 1 'งานใหม่
                subStr = " ((isPrinted<>'Y' Or isPrinted is null) and (isCanceled<>'Y' Or isCanceled is null) and (isClosed<>'Y' Or isClosed is null)) "
            Case 2 'พิมพ์ Menifest แล้ว
                subStr = " (isPrinted='Y' and (isCanceled<>'Y' Or isCanceled is null) and (isClosed<>'Y' Or isClosed is null)) "
            Case 3 'ปิดงาน
                subStr = " (isClosed='Y' and (ISNULL(isCanceled, 'N') <>'Y')) "
                'subStr = " (isClosed='Y' and (isCanceled <>'Y' Or isCanceled is null)) "
            Case 4 'ยกเลิก
                subStr = " (isCanceled='Y') "
        End Select
        If cboJobStatus.ListIndex > 0 Then
            If InStr(1, UCase(strCond), "WHERE") > 0 Then
                strCond = strCond & " AND " & subStr
            Else
                strCond = strCond & " WHERE " & subStr
            End If
        End If
        
    End If
    curCondition_ = strCond
    sqlStr = sqlStr & strCond
    sqlStr = SelectWithCompanyID(sqlStr, False)
        
    sqlStr = sqlStr & " ORDER BY WorkDate, MenifestNo, DocNo"
    Set rs = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    
    If rs Is Nothing Then
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ"
        Exit Sub
    End If
    
    If rs.RecordCount = 0 Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    frmMain.Stb1.Panels(1).Text = "Calculating ..."
    frmMain.Stb1.Panels(1).Text = "Menifest Count: " & rs.RecordCount
    
    '' ******* Cal PreInvoice *********************
    Dim tmpS As New BWGSearchManager, tmpItem As BWGMenifest, tmpJobDet As BWGJobDataDetail
    Dim selMenifestID As String, tmpTimeTableNo As String, LastTimeTable As Collection
    Dim colTrip As Collection, colMinWeightCheck As New Collection, tmpWaste As BWGWasteDataCR, company_ As BWGCompanyConfig
    Dim tmpCol As Collection, tmpCol2 As Collection
    Dim curTreatUnitID As String, curMfID As String, maxType As String, minType As String, tmpTotalAmt As Double
    Dim tmpTranUnit As String, nonGroupWeight As Double, maxNonTreatRate As Double, maxNonUnit$, curGroupWeight As Double, tmpWeight As Double
    Dim tmpWasteType As String, curTripID As String
    Dim curUnitTimeTable As String, ExtraCharge As String
    Dim z%, X1%, curRow%, i As Integer
    Dim TWeight As Double
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    selMenifestID = ""
    Do While Not rs.EOF
        If Trim(rs!menifestID) <> "" Then
            selMenifestID = selMenifestID & "'" & rs!menifestID & "',"
        End If
        rs.MoveNext
    Loop
    
    If InStr(1, selMenifestID, ",") > 0 Then
        selMenifestID = Left(selMenifestID, Len(selMenifestID) - 1)
    End If
    
    If Trim(selMenifestID) = "" Then Exit Sub
                        
    CalMinMaxWeight selMenifestID, colTrip, colMinWeightCheck

    frmMain.Stb1.Panels(2).Text = ""
    flexItems.Clear
    
    mDocTitle = "บริษัท :  " & cboCompany.Text
    mDocTitleCust = "บริษัทลูกค้า  : " & curCustomer_.CustomerName
    mBackColor = &O0
    SetTitleGrid
    SetCustGrid
    SetItemsGrid
    curRow = 1
    With flexItems
        .Rows = 3
        curTreatUnitID = ""
        Set tmpCol2 = tmpS.MenifestSearch("MenifestID in (" & selMenifestID & ") And (isCanceled<>'Y' Or isCanceled is null)")
        frmMain.Stb1.Panels(2).Text = "Get Manifest Count " & tmpCol2.Count
        For x = 1 To tmpCol2.Count
            Set tmpItem = tmpCol2(x)
            If Trim(curTripID) <> tmpItem.TimeTableNo & "_" & tmpItem.TripNo Then
                If Trim(curTripID) = "" Then
                    nonGroupWeight = 0
                    curGroupWeight = 0
                ElseIf tmpItem.TimeTableNo <> Replace(Trim(Split(curTripID, "_")(0)), "'", "") Then
                    nonGroupWeight = 0
                    curGroupWeight = 0
                End If
                curTreatUnitID = ""
                curTripID = tmpItem.TimeTableNo & "_" & tmpItem.TripNo
            End If
            If Trim(tmpItem.ID) <> Trim(curMfID) Then
                curMfID = Trim(tmpItem.ID)
                If maxType < LCase(tmpItem.NonOrHz) Then maxType = LCase(tmpItem.NonOrHz)
                If minType > LCase(tmpItem.NonOrHz) Then minType = LCase(tmpItem.NonOrHz)
                ''''************************ ADD TimeTableNo IN HEADER **************************
'                If Option2.Value = True Then
'                    If curUnitTimeTable = "" Or curUnitTimeTable <> tmpItem.TimeTableNo Then
'                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
'                        .TextMatrix(curRow, 8) = tmpItem.ID
'                        .TextMatrix(curRow, 1) = "ใบคุมเลขที่ " & tmpItem.TimeTableNo
'                        If tmpItem.TripMoreWorkerFee > 0 Then
'                         .TextMatrix(curRow, 1) = .TextMatrix(curRow, 1) & "  (" & "ค่าแรงคนงาน" & ")"
'                        .TextMatrix(curRow, 7) = tmpItem.TripMoreWorkerFee
'                        End If
'                       ''.Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
'                       ''.Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
'                        curRow = curRow + 1
'                        .Rows = .Rows + 1
'                    End If
'                End If
                ''''********************************************************************************************
                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                .TextMatrix(curRow, 1) = "วันที่ " & FormatYMD_to_DMY(tmpItem.WorkDate, "/", "/") & " เลขที่ " & tmpItem.DocNo
                .TextMatrix(curRow, 2) = ""
                .TextMatrix(curRow, 3) = ""
                .TextMatrix(curRow, 4) = ""
                .TextMatrix(curRow, 5) = ""
                .TextMatrix(curRow, 6) = ""
                .TextMatrix(curRow, 7) = ""
                .TextMatrix(curRow, 8) = tmpItem.ID
                .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                curRow = curRow + 1
                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                If Trim(tmpItem.InOtherDesc) <> "" Then
                    .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                    .TextMatrix(curRow, 1) = "     " & "     " & tmpItem.InOtherDesc
                    .TextMatrix(curRow, 2) = ""
                    .TextMatrix(curRow, 3) = "1"
                    .TextMatrix(curRow, 4) = ""
                    .TextMatrix(curRow, 5) = tmpItem.InOtherFee
                    .TextMatrix(curRow, 6) = ""
                    .TextMatrix(curRow, 7) = tmpItem.InOtherFee
                    .TextMatrix(curRow, 8) = tmpItem.ID
                    tmpTotalAmt = tmpTotalAmt + tmpItem.InOtherFee
                    curRow = curRow + 1
                    If .Rows - 1 = curRow Then .Rows = .Rows + 5
                End If
                
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Set tmpCol = Nothing
            Set tmpCol = New Collection
            For z = 1 To colTrip.Count
                If Trim(colTrip(z).menifestID) = Trim(tmpItem.ID) Then
                    tmpCol.Add colTrip(z), "'" & colTrip(z).ID & "'"
                End If
            Next
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Set tmpCol = tmpS.JobDetailInTripSearch("MenifestID='" & tmpItem.ID & "'", "WasteType,TreatmentUnitID")
            If tmpCol.Count <= 0 Then
                frmMain.Stb1.Panels(1).Text = "Menifest detail can't not load, system will be load data again."
                frmMain.Stb1.Panels(2).Text = "Menifest detail: " & Trim("" & tmpCol.Count)
            End If
            For y = 1 To tmpCol.Count
                Set tmpJobDet = tmpCol(y)
                If curUnitTimeTable <> tmpJobDet.TimeTableNo Then
                    If curUnitTimeTable <> "" Then
                        maxNonTreatRate = 0: curGroupWeight = 0: nonGroupWeight = 0
                    End If
                End If
                tmpTranUnit = tmpJobDet.QuoTransUnitID
'                If curTreatUnitID = "" Then curTreatUnitID = Trim(tmpJobDet.QuoTreatmentUnitID)
                If Trim(tmpJobDet.FailType) = "" Then
                    .TextMatrix(curRow, 8) = tmpItem.ID
                    .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                    .TextMatrix(curRow, 1) = "     " & "     " & tmpJobDet.WasteName
                    .TextMatrix(curRow, 2) = tmpJobDet.WasteQuotationNo & " : " & tmpJobDet.QuoTreatmentRate & " " & tmpJobDet.QuoTreatmentUnitID & "{" & tmpItem.CustomerName & "}"
                    If UCase(tmpJobDet.WasteType) = "NON" Or UCase(tmpJobDet.WasteType) = "NBL" Then
                        If tmpJobDet.ChargeWeightType = "A" Then
                            nonGroupWeight = nonGroupWeight + tmpJobDet.DisposerWeight
                        Else
                            nonGroupWeight = nonGroupWeight + tmpJobDet.EstWasteQty
                        End If
                        If maxNonTreatRate <= tmpJobDet.JobTreatmentRate Then
                            maxNonTreatRate = tmpJobDet.JobTreatmentRate
                            maxNonUnit = tmpJobDet.QuoTreatmentUnitID
                        End If
                    End If
                    If tmpJobDet.QuoTreatmentUnitID = "บาท/เที่ยว" Then
                        If Trim(curTreatUnitID) = "" And curUnitTimeTable <> tmpJobDet.TimeTableNo Then
                            .TextMatrix(curRow - 1, 2) = ""
                            .TextMatrix(curRow - 1, 3) = "1"
                            .TextMatrix(curRow - 1, 4) = "เที่ยว"
                            .TextMatrix(curRow - 1, 5) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                            .TextMatrix(curRow - 1, 6) = tmpJobDet.QuoTreatmentUnitID
                            .TextMatrix(curRow - 1, 7) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + tmpJobDet.JobTreatmentRate
                            curUnitTimeTable = tmpJobDet.TimeTableNo
                        End If
                        If UCase(tmpItem.NonOrHz) = "NON" Or UCase(tmpItem.NonOrHz) = "NBL" Then
                            .TextMatrix(curRow, 2) = ""
                            .TextMatrix(curRow, 3) = ""
                            .TextMatrix(curRow, 4) = ""
                            .TextMatrix(curRow, 5) = ""
                            .TextMatrix(curRow, 6) = ""
                            .TextMatrix(curRow, 7) = ""
                        Else
                            .TextMatrix(curRow, 2) = ""
                            If tmpJobDet.ChargeWeightType = "A" Then
                                .TextMatrix(curRow, 3) = Format(tmpJobDet.DisposerWeight, "#,##0.000")
                            Else
                                .TextMatrix(curRow, 3) = Format(tmpJobDet.EstWasteQty, "#,##0.000")
                            End If
                            .TextMatrix(curRow, 4) = "ตัน"
                            .TextMatrix(curRow, 5) = ""
                            .TextMatrix(curRow, 6) = ""
                            .TextMatrix(curRow, 7) = ""
                        End If
                        If tmpJobDet.ChargeWeightType = "A" Then
                            curGroupWeight = curGroupWeight + tmpJobDet.DisposerWeight
                        Else
                            curGroupWeight = curGroupWeight + tmpJobDet.EstWasteQty
                        End If
                        curTreatUnitID = Trim(tmpJobDet.QuoTreatmentUnitID)
                    Else
                        If Trim(curTreatUnitID) <> "" Then
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "     " & "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = ""
                            .TextMatrix(curRow, 3) = Format(tmpJobDet.MenifestActQty, "#,##0.000")
                            .TextMatrix(curRow, 4) = "ตัน"
                            .TextMatrix(curRow, 5) = ""
                            .TextMatrix(curRow, 6) = ""
                            .TextMatrix(curRow, 7) = ""
                            .TextMatrix(curRow, 8) = tmpItem.ID
                            .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                            .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            
                  ' **** ค่าแรงคนงาน *******
'                            If tmpItem.TripMoreWorkerFee > 0 Then
'                                .TextMatrix(curRow, 8) = tmpItem.ID
'                                .TextMatrix(curRow, 1) = "ค่าแรงคนงาน"
'                                .TextMatrix(curRow, 7) = tmpItem.TripMoreWorkerFee
'                                curRow = curRow + 1
'                                .Rows = .Rows + 1
'                            End If
                    '**************************
                   
                        End If
                        curTreatUnitID = ""
                        curUnitTimeTable = tmpJobDet.TimeTableNo
                        curGroupWeight = 0
                        If UCase(tmpJobDet.WasteType) <> "NON" And UCase(tmpJobDet.WasteType) <> "NBL" Then
                            .TextMatrix(curRow, 2) = ""
                            If tmpJobDet.ChargeWeightType = "A" Then
                                .TextMatrix(curRow, 3) = Format(tmpJobDet.DisposerWeight, "#,##0.000")
                                .TextMatrix(curRow, 4) = Replace(tmpJobDet.QuoTreatmentUnitID, "บาท/", "")
                                tmpWeight = tmpJobDet.DisposerWeight
                            Else
                                .TextMatrix(curRow, 3) = Format(tmpJobDet.EstWasteQty, "#,##0.000")
                                .TextMatrix(curRow, 4) = Replace(tmpJobDet.QuoTreatmentUnitID, "บาท/", "")
                                tmpWeight = tmpJobDet.EstWasteQty
                            End If
                            If colMinWeightCheck("'" & tmpJobDet.TimeTableNo & "'").isUnderWeight And LCase(Trim(colMinWeightCheck("'" & tmpJobDet.TimeTableNo & "'").MaxWasteID)) = LCase(Trim(tmpJobDet.wastedataID)) Then
                                If LCase(tmpJobDet.WasteType) = "haz" Then
                                    If tmpJobDet.HazMinRate > 0 Then
                                        .TextMatrix(curRow, 5) = Format(tmpJobDet.HazMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 5) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                                    End If
                                ElseIf LCase(tmpJobDet.WasteType) = "hbl" Then
                                    If tmpJobDet.HBLMinRate > 0 Then
                                        .TextMatrix(curRow, 5) = Format(tmpJobDet.HBLMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 5) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                                    End If
                                Else
                                    If tmpJobDet.HazMinRate > 0 Then
                                        .TextMatrix(curRow, 5) = Format(tmpJobDet.HazMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 5) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                                    End If
                                End If
                                .row = curRow
                                .col = 5
                                .CellForeColor = vbRed
                            Else
                                .TextMatrix(curRow, 5) = Format(tmpJobDet.JobTreatmentRate, "#,##0.00")
                            End If
                            .TextMatrix(curRow, 6) = tmpJobDet.QuoTreatmentUnitID
                            .TextMatrix(curRow, 7) = Format(tmpWeight * CDbl(.TextMatrix(curRow, 5)), "#,##0.00")
                            tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 7))
                        End If
                    End If
                    curRow = curRow + 1
                    If .Rows - 1 = curRow Then .Rows = .Rows + 5
                End If
                tmpWasteType = tmpJobDet.WasteType
                tmpTimeTableNo = tmpJobDet.TimeTableNo
                Set tmpJobDet = Nothing
            Next y
            
            Select Case UCase(tmpWasteType)
                Case "NON", "NBL"
                    If InStr(1, maxNonUnit, "ตัน") > 0 Then
                        If colMinWeightCheck("'" & tmpTimeTableNo & "'").isUnderWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "     " & "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = ""
                            .TextMatrix(curRow, 3) = Format(nonGroupWeight, "#,##0.000")
                            .TextMatrix(curRow, 4) = "ตัน"
                            .TextMatrix(curRow, 5) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 6) = maxNonUnit
                            .TextMatrix(curRow, 7) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 8) = tmpItem.ID
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                            .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.ID) Then
                                Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                                .TextMatrix(curRow, 8) = tmpItem.ID
                                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                ExtraCharge = "     " & "น้ำหนักไม่ถึง " & tmpWaste.MinWeightPerCar & " ตัน คิดน้ำหนักเพิ่ม"
                                .TextMatrix(curRow, 1) = ExtraCharge
                                .TextMatrix(curRow, 2) = ""
        '                        .TextMatrix(curRow, 2) = Format(tmpWaste.MinWeightPerCar, "#,##0.000")
                                '===============================
                                .TextMatrix(curRow, 3) = Format(tmpWaste.MinWeightPerCar - colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight, "#,##0.000")
                                TWeight = CDbl(.TextMatrix(curRow, 3))
                                '===============================
                                .TextMatrix(curRow, 4) = "ตัน"
                                If LCase(tmpWaste.WasteType) = "non" Then
                                    If tmpWaste.NonMinRate > 0 Then
                                        .TextMatrix(curRow, 5) = Format(tmpWaste.NonMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 5) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                    End If
                                ElseIf LCase(tmpWaste.WasteType) = "nbl" Then
                                    If tmpWaste.NBLMinRate > 0 Then
                                        .TextMatrix(curRow, 5) = Format(tmpWaste.NBLMinRate, "#,##0.00")
                                    Else
                                        .TextMatrix(curRow, 5) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                    End If
                                End If
                                .TextMatrix(curRow, 6) = "บาท/ตัน"
        '                        .TextMatrix(curRow, 6) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 2)), "#,##0.00")
                                '===============
                                .TextMatrix(curRow, 7) = Format(CDbl(.TextMatrix(curRow, 5)) * TWeight, "#,##0.00")
                                '===============
                                tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 7))
                                .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                                .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                                '=================================
                                curRow = curRow + 1
                                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                
                                   ' **** ค่าแรงคนงาน *******
'                                If tmpItem.TripMoreWorkerFee > 0 Then
'                                    .TextMatrix(curRow, 8) = tmpItem.ID
'                                    .TextMatrix(curRow, 1) = "ค่าแรงคนงาน"
'                                    .TextMatrix(curRow, 7) = tmpItem.TripMoreWorkerFee
'                                    curRow = curRow + 1
'                                    .Rows = .Rows + 1
'                                End If
                                         '**************************
                                
                            End If
                        ElseIf colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                            
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "     " & "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = ""
                            .TextMatrix(curRow, 3) = Format(nonGroupWeight, "#,##0.000")
                            .TextMatrix(curRow, 4) = "ตัน"
                            '=================================
                            .TextMatrix(curRow, 5) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 6) = maxNonUnit
                            .TextMatrix(curRow, 7) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 8) = tmpItem.ID
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                            .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.ID) Then
                                Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                                
                                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                .TextMatrix(curRow, 1) = "     " & "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " ตัน คิดน้ำหนักส่วนเกิน"
                                .TextMatrix(curRow, 2) = ""
                                .TextMatrix(curRow, 3) = Format(colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar, "#,##0.000")
                                .TextMatrix(curRow, 4) = "ตัน"
                                .TextMatrix(curRow, 5) = Format(tmpWaste.MaxRate, "#,##0.00")
                                .TextMatrix(curRow, 6) = tmpWaste.MaxRateType
                                If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                    .TextMatrix(curRow, 7) = Format(tmpWaste.MaxRate, "#,##0.00")
                                Else
                                    .TextMatrix(curRow, 7) = Format(CDbl(.TextMatrix(curRow, 4)) * CDbl(.TextMatrix(curRow, 3)), "#,##0.00")
                                End If
                                .TextMatrix(curRow, 8) = tmpItem.ID
                                tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 7))
                                .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                                .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                                '=================================
                                curRow = curRow + 1
                                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                   ' **** ค่าแรงคนงาน *******
                                 If tmpItem.TripMoreWorkerFee > 0 Then
                                    .TextMatrix(curRow, 8) = tmpItem.ID
                                    .TextMatrix(curRow, 1) = "ค่าแรงคนงาน"
                                    .TextMatrix(curRow, 7) = tmpItem.TripMoreWorkerFee
                                    curRow = curRow + 1
                                    .Rows = .Rows + 1
                                End If
                                '**************************
                            End If
                        Else
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "     " & "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = ""
                            .TextMatrix(curRow, 3) = Format(nonGroupWeight, "#,##0.000")
                            .TextMatrix(curRow, 4) = "ตัน"
                            .TextMatrix(curRow, 5) = Format(maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 6) = maxNonUnit
                            .TextMatrix(curRow, 7) = Format(nonGroupWeight * maxNonTreatRate, "#,##0.00")
                            .TextMatrix(curRow, 8) = tmpItem.ID
                            tmpTotalAmt = tmpTotalAmt + Format(nonGroupWeight * maxNonTreatRate, "0.00") * 1
                            .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                            .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                               ' **** ค่าแรงคนงาน *******
                            If tmpItem.TripMoreWorkerFee > 0 Then
                                .TextMatrix(curRow, 8) = tmpItem.ID
                                .TextMatrix(curRow, 1) = "ค่าแรงคนงาน"
                                .TextMatrix(curRow, 7) = tmpItem.TripMoreWorkerFee
                                curRow = curRow + 1
                                .Rows = .Rows + 1
                            End If
                    '**************************
                            
                        End If
                        nonGroupWeight = 0
                        maxNonTreatRate = 0
                    Else
                        GoTo RatePerTrip
                    End If
                    
                Case "HAZ", "HBL", "HIC", "HBI", "BBH", "BBL", "NIC", "NBI"
                    If InStr(1, Trim(curTreatUnitID), "เที่ยว") = 0 Then
                        If Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID) <> "" Then
                            Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                            If InStr(1, tmpWaste.TreatmentUnitID, "ตัน") > 0 Then
                                If LCase(Trim(tmpWaste.WasteType)) = LCase(Trim(tmpWasteType)) And Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").menifestID) = Trim(tmpItem.ID) Then
                                    If colMinWeightCheck("'" & tmpTimeTableNo & "'").isUnderWeight Then
                                        .TextMatrix(curRow, 8) = tmpItem.ID
                                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                        If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                            curRow = curRow + 1
                                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                        End If
                                        .TextMatrix(curRow, 1) = "     " & "น้ำหนักไม่ถึง " & tmpWaste.MinWeightPerCar & " ตัน คิดน้ำหนักเพิ่มอีก"
                                        .TextMatrix(curRow, 2) = ""
                                        .TextMatrix(curRow, 3) = Format(tmpWaste.MinWeightPerCar - colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight, "#,##0.000")
                                        .TextMatrix(curRow, 4) = "ตัน"
                                        If LCase(tmpWaste.WasteType) = "haz" Then
                                            If tmpWaste.HazMinRate > 0 Then
                                                .TextMatrix(curRow, 5) = Format(tmpWaste.HazMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 5) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        ElseIf LCase(tmpWaste.WasteType) = "hbl" Then
                                            If tmpWaste.HBLMinRate > 0 Then
                                                .TextMatrix(curRow, 5) = Format(tmpWaste.HBLMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 5) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        Else
                                            If tmpWaste.HazMinRate > 0 Then
                                                .TextMatrix(curRow, 5) = Format(tmpWaste.HazMinRate, "#,##0.00")
                                            Else
                                                .TextMatrix(curRow, 5) = Format(tmpWaste.TreatmentFee, "#,##0.00")
                                            End If
                                        End If
                                        .TextMatrix(curRow, 6) = "บาท/ตัน"
                                        .TextMatrix(curRow, 7) = Format(CDbl(.TextMatrix(curRow, 5)) * CDbl(.TextMatrix(curRow, 3)), "#,##0.00")
                                        tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 7))
                                        .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                                        .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                                        '=================================
                                        curRow = curRow + 1
                                        If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                    ElseIf colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight Then
                                        .TextMatrix(curRow, 8) = tmpItem.ID
                                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                        If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                            curRow = curRow + 1
                                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                        End If
                                        .TextMatrix(curRow, 1) = "     " & "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " ตัน คิดน้ำหนักส่วนเกิน"
                                        .TextMatrix(curRow, 2) = ""
                                        .TextMatrix(curRow, 3) = Format(colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar, "#,##0.000")
                                        .TextMatrix(curRow, 4) = "ตัน"
                                        .TextMatrix(curRow, 5) = Format(tmpWaste.MaxRate, "#,##0.00")
                                        .TextMatrix(curRow, 6) = tmpWaste.MaxRateType
                                        If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                            .TextMatrix(curRow, 7) = Format(tmpWaste.MaxRate, "#,##0.00")
                                        Else
                                            .TextMatrix(curRow, 7) = Format(CDbl(.TextMatrix(curRow, 5)) * CDbl(.TextMatrix(curRow, 3)), "#,##0.00")
                                        End If
                                        tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 7))
                                        .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                                        .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                                        '=================================
                                        curRow = curRow + 1
                                        If .Rows - 1 = curRow Then .Rows = .Rows + 5
                                    End If
                                End If
                            End If
                            Set tmpWaste = Nothing
                        Else
                            GoTo RatePerTrip
                        End If
                    Else
                        GoTo RatePerTrip
                    End If
                    
                Case Else
RatePerTrip:
                    If Trim(curTreatUnitID) <> "" Then
                        If colMinWeightCheck("'" & tmpTimeTableNo & "'").isOverWeight And LCase(Trim(colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteType)) = LCase(Trim(tmpWasteType)) Then
                            Set tmpWaste = colMinWeightCheck("'" & tmpTimeTableNo & "'").ColWasteWeight("'" & colMinWeightCheck("'" & tmpTimeTableNo & "'").MaxWasteID & "'")
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "     " & "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = ""
                            .TextMatrix(curRow, 3) = Format(curGroupWeight, "#,##0.000")
                            .TextMatrix(curRow, 4) = "ตัน"
                            .TextMatrix(curRow, 5) = ""
                            .TextMatrix(curRow, 6) = ""
                            .TextMatrix(curRow, 7) = ""
                            .TextMatrix(curRow, 8) = tmpItem.ID
                            .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                            .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                            '=================================
                            If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                curRow = curRow + 1
                                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            End If
                            If Trim(.TextMatrix(curRow, 1)) <> "" Then
                                curRow = curRow + 1
                                If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            End If
                            If tmpWaste.menifestID = tmpItem.ID Then
                                .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                                .TextMatrix(curRow, 1) = "     " & "น้ำหนักเกินที่กำหนด " & tmpWaste.MaxWeightPerCar & " ตัน คิดน้ำหนักส่วนเกิน"
                                .TextMatrix(curRow, 2) = ""
                                .TextMatrix(curRow, 3) = Format(colMinWeightCheck("'" & tmpTimeTableNo & "'").TotalWeight - tmpWaste.MaxWeightPerCar, "#,##0.000")
                                .TextMatrix(curRow, 4) = "ตัน"
                                .TextMatrix(curRow, 5) = Format(tmpWaste.MaxRate, "#,##0.00")
                                .TextMatrix(curRow, 6) = tmpWaste.MaxRateType
                                If InStr(1, tmpWaste.MaxRateType, "เที่ยว") > 0 Then
                                    .TextMatrix(curRow, 7) = Format(tmpWaste.MaxRate, "#,##0.00")
                                Else
                                    .TextMatrix(curRow, 7) = Format(CDbl(.TextMatrix(curRow, 5)) * CDbl(.TextMatrix(curRow, 3)), "#,##0.00")
                                End If
                                .TextMatrix(curRow, 8) = tmpItem.ID
                                tmpTotalAmt = tmpTotalAmt + CDbl(.TextMatrix(curRow, 7))
                            End If
                            .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                            .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                        Else
                            .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                            .TextMatrix(curRow, 1) = "     " & "น้ำหนักรวม"
                            .TextMatrix(curRow, 2) = ""
                            .TextMatrix(curRow, 3) = Format(curGroupWeight, "#,##0.000")
                            .TextMatrix(curRow, 4) = "ตัน"
                            .TextMatrix(curRow, 5) = ""
                            .TextMatrix(curRow, 6) = ""
                            .TextMatrix(curRow, 7) = ""
                            .TextMatrix(curRow, 8) = tmpItem.ID
                            .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                            .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                            '=================================
                            curRow = curRow + 1
                            If .Rows - 1 = curRow Then .Rows = .Rows + 5
                            curGroupWeight = 0
                        End If
                    End If
            End Select
            
            '''' **************** ADDING TRANSPORT FEE *************************
            Dim xTrip As Integer, LMenifest As Integer
            LMenifest = tmpItem.getMenifestHeader.Count
            For xTrip = 1 To tmpItem.getMenifestHeader.Count
                If tmpItem.TripNo = 1 And tmpItem.getMenifestHeader(xTrip).TripNo = 1 Then
                    CarRegisID1 = tmpItem.CarRegisID
                ElseIf tmpItem.TripNo = 2 And tmpItem.getMenifestHeader(xTrip).TripNo = 2 Then
                    CarRegisID2 = tmpItem.CarRegisID
                End If
                If tmpItem.DocNo = tmpItem.getMenifestHeader(LMenifest).DocNo Then
                    If tmpItem.getMenifestHeader(xTrip).isPriceIncTrans = False Then
                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                        .TextMatrix(curRow, 8) = tmpItem.ID
                        If tmpItem.JobDetails(1).isTrailer = True Then
                            .TextMatrix(curRow, 1) = "     " & "ค่าขนส่ง " & tmpItem.TruckTypeName & " ( " & CarRegisID1 & ", " & CarRegisID2 & ")"
                        Else
                            .TextMatrix(curRow, 1) = "     " & "ค่าขนส่ง " & tmpItem.TruckTypeName & " ( " & tmpItem.CarRegisID & ")"
                        End If
                        .TextMatrix(curRow, 2) = ""
                        .TextMatrix(curRow, 3) = "1"
                        .TextMatrix(curRow, 4) = "เที่ยว"
                        .TextMatrix(curRow, 5) = Format(tmpItem.JobDetails(1).QuoTransportFee, "#,##0.00")
                        .TextMatrix(curRow, 6) = tmpItem.JobDetails(1).QuoTransUnitID
                        .TextMatrix(curRow, 7) = Format(tmpItem.JobDetails(1).QuoTransportFee, "#,##0.00")
        
                        .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                        .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                        curRow = curRow + 1
                        .Rows = .Rows + 1
                  ' **** ค่าแรงคนงาน *******
                      If tmpItem.TripMoreWorkerFee > 0 Then
                            .TextMatrix(curRow, 8) = tmpItem.ID
                            .TextMatrix(curRow, 1) = "ค่าแรงคนงาน"
                            .TextMatrix(curRow, 7) = tmpItem.TripMoreWorkerFee
                            curRow = curRow + 1
                            .Rows = .Rows + 1
                         End If
                    '**************************
                        
                        
                        Exit For
                    Else
                        .TextMatrix(curRow, 0) = tmpItem.JobDataCarID & "_" & tmpItem.TripNo
                        .TextMatrix(curRow, 8) = tmpItem.ID
                        tmpItem.getTripControl
                        If tmpItem.JobDetails(1).isTrailer = True Then
                                .TextMatrix(curRow, 1) = "     " & "รวมค่าขนส่ง " & tmpItem.TruckTypeName & " ( " & tmpItem.CarRegisID & ", " & CarRegisID2 & ")"
                        Else
                            .TextMatrix(curRow, 1) = "     " & "รวมค่าขนส่ง " & tmpItem.TruckTypeName & " ( " & tmpItem.CarRegisID & ")"
                        End If
                        .TextMatrix(curRow, 2) = ""
                        .TextMatrix(curRow, 3) = "1"
                        .TextMatrix(curRow, 4) = "เที่ยว"
                        '''******* แสดงค่าขนส่งกรณีรวมค่าขนส่งเฉพาะ AKP ***************
                        If tmpItem.TSDFProvinceName = "สมุทรปราการ" Then
                            .TextMatrix(curRow, 5) = Format(tmpItem.TripTransFeeInc, "#,##0.00")
                            .TextMatrix(curRow, 6) = "บาท/เที่ยว"
                            .TextMatrix(curRow, 7) = Format(tmpItem.TripTransFeeInc, "#,##0.00")
                        Else
                            .TextMatrix(curRow, 5) = ""
                            .TextMatrix(curRow, 6) = ""
                            .TextMatrix(curRow, 7) = ""
                        End If
                        '''******* End แสดงค่าขนส่งกรณีรวมค่าขนส่งเฉพาะ AKP *********
                        .Cell(flexcpFontBold, curRow, 0, curRow, .Cols - 1) = mBackColor
                        .Cell(flexcpForeColor, curRow, 0, curRow, .Cols - 1) = mBackColor
                        curRow = curRow + 1
                        .Rows = .Rows + 1
                        
                           ' **** ค่าแรงคนงาน *******
                        If tmpItem.TripMoreWorkerFee > 0 Then
                            .TextMatrix(curRow, 8) = tmpItem.ID
                            .TextMatrix(curRow, 1) = "ค่าแรงคนงาน"
                            .TextMatrix(curRow, 7) = tmpItem.TripMoreWorkerFee
                            curRow = curRow + 1
                            .Rows = .Rows + 1
                         End If
                         '**************************
                
                        Exit For
                    End If
      
                End If
            Next xTrip
            Set tmpItem = Nothing
        Next x
        
        
        ' ******************รวมค่าบริการ ทั้งหมด*******************
          Dim sumAmunt As Double
          Dim j As Integer
            For j = 1 To curRow
                If .TextMatrix(j, 7) <> "" Then
                    sumAmunt = sumAmunt + CDbl(.TextMatrix(j, 7))
                End If
             Next
            curRow = curRow + 2
            .Rows = .Rows + 2
             .TextMatrix(curRow, 6) = "รวม    "
             .TextMatrix(curRow, 7) = Format(sumAmunt, "#,##0")
             curRow = curRow + 1
            .Rows = .Rows + 1
             .TextMatrix(curRow, 6) = "VAT 7 %    "
             Dim VaT As Double
             VaT = (sumAmunt * 7) / 100
             .TextMatrix(curRow, 7) = Format(VaT, "#,##0")
            curRow = curRow + 1
            .Rows = .Rows + 1
             Dim sum As Double
             sum = sumAmunt + VaT
            .TextMatrix(curRow, 6) = "รวมสุทธิ    "
            .TextMatrix(curRow, 7) = Format(sum, "#,##0")
            curRow = curRow + 1
            .Rows = .Rows + 1
            .TextMatrix(curRow, 1) = "หมายเหตุ "
            curRow = curRow + 1
            .Rows = .Rows + 1
            .TextMatrix(curRow, 1) = "1.เอกสารฉบับนี้ไม่ใช่ใบเสนอราคา/ใบแจ้งค่าบริการ"
             curRow = curRow + 1
            .Rows = .Rows + 1
            .TextMatrix(curRow, 1) = "2.กรุณาส่งใบสั่งซื้อ(PO) มาที่ dada.akp@hotmail.com"
              curRow = curRow + 1
            .Rows = .Rows + 1
            .TextMatrix(curRow, 1) = "   หรือ เบอร์แฟกซ์  02-3230724"
             curRow = curRow + 1
            .Rows = .Rows + 1
            .TextMatrix(curRow, 1) = "   สอบถามได้ที่ โทร 02-3230714-19 ต่อ 153 คุณด้า"
            curRow = curRow + 1
            .Rows = .Rows + 1
        '  ******************************************************

    End With
    
    
    Screen.MousePointer = 0
    MsgBox "พบข้อมูล " & Format(rs.RecordCount, "#,##0") & " รายการ"
    frmMain.Stb1.Panels(1).Text = "พบข้อมูล " & Format(rs.RecordCount, "#,##0") & " รายการ"
    rs.Close
    Set rs = Nothing
    Exit Sub
    
ErrD:
    Screen.MousePointer = 0
    GetMsg "", vbCritical
    Exit Sub
    
End Sub

Private Sub Command1_Click()

End Sub


Private Sub flexItems_AfterUserResize(ByVal row As Long, ByVal col As Long)
    flexTitle.ColWidth(0) = flexItems.ColWidth(0) + flexItems.ColWidth(1) + flexItems.ColWidth(2) + flexItems.ColWidth(3)
    flexTitle.ColWidth(1) = flexItems.ColWidth(4) + flexItems.ColWidth(5)
    flexTitle.ColWidth(3) = flexItems.ColWidth(6) + flexItems.ColWidth(7) - flexTitle.ColWidth(2)
End Sub

Private Sub Form_Load()
    With cboJobStatus
        .AddItem "== ทุกสถานะ =="
        .AddItem "งานใหม่"
        .AddItem "พิมพ์ Manifest แล้ว"
        .AddItem "ปิดงานแล้ว"
        .AddItem "ยกเลิก"
        .ListIndex = 0
    End With
    PopulateCompany cboCompany, cboIndex
    mDocTitle = "PRE INVOICE"
    mBackColor = &HFFFFFF     ' Blue
    SetTitleGrid
    SetItemsGrid
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Picture1.Height = Me.Height - 60
    Picture2.Height = Me.Height - 60
    Picture2.Width = Me.Width - Picture1.Width - 150
    flexTitle.Width = Picture2.Width
    flexItems.Width = Picture2.Width - 150
    flexItems.Height = Picture2.Height - flexItems.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VisibleImage True
End Sub



Private Sub TxtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCustomer_Validate False
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    Dim selStr As String
    If Trim(txtCustomer.Text) = "" Then Exit Sub
    If Trim(txtCustomer.Text) = "*" Then Exit Sub
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    selStr = Trim(curCustomer_.CustomerName)
    If Trim(txtCustomer.Text) = "" And Trim(selStr) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(selStr) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .ShowSubContact = True
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .GetCriteriaOptonal ""
'                .GetCriteriaOptonal " AND CustGroupID <> '488F2E5C057F46869D1A56185B11D644'"
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_.ID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Private Sub txtDocNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdSearch_Click
End Sub

Private Sub txtJobNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If Trim(txtSale.Text) = "" Then Exit Sub
    If Trim(txtSale.Text) = "*" Then Exit Sub
    If SaleStaff_ Is Nothing Then Set SaleStaff_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(SaleStaff_.StaffFName & " " & SaleStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set SaleStaff_ = Nothing
        Set SaleStaff_ = New BWGCompanyStaff
    Else
        If Trim(SaleStaff_.StaffFName & " " & SaleStaff_.StaffLName) <> Trim(txtSale.Text) Then
            Set SaleStaff_ = Nothing
            Set SaleStaff_ = New BWGCompanyStaff
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtSale.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If SaleStaff_.ID = "" Then txtSale.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCust As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curCustomer_ = selCust
    With curCustomer_
        txtCustomer.Text = .CustomerName
        If curCustomer_.ID = "" Then txtCustomer.Text = ""
    End With
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set SaleStaff_ = Nothing
    Set SaleStaff_ = selStaff
    If SaleStaff_ Is Nothing Then Exit Sub
    With SaleStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub

Private Sub PrintTitle(ByRef flex As VSFlexGrid, mText$, mRow&, mCol&)

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

Private Sub SetTitleGrid()
    
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
        .AllowUserResizing = flexResizeColumns
        .Editable = flexEDKbdMouse
        ' set columns width
        .ColWidth(0) = 6000
        .ColWidth(1) = 0
        .ColWidth(2) = 50
        .ColWidth(3) = 0
        ' Place title
        .MergeCol(0) = True
        .Cell(flexcpText, 0, 0, 1, 0) = mDocTitle
        .Cell(flexcpFontName, 0, 0, 1, 0) = "MS Sans Serif"
'        .Cell(flexcpFontName, 0, 0, 1, 0) = "Haettenschweiler"
        .Cell(flexcpFontSize, 0, 0, 1, 0) = "16"
        .Cell(flexcpForeColor, 0, 0, 1, 0) = mBackColor
        
        ' invoice No.
        PrintTitle flexTitle, "CUSTOMER", 1, 1
        PrintTitle flexTitle, "PRE INVOICE NO.", 1, 3
        
        ' draw borders
        .Select 1, 3, 2, 3
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select 1, 1, 2, 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        
        'set column aligment
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        
        .Redraw = True
        
    End With
    
End Sub

Private Sub SetCustGrid()
    
    ' FlexTitle
    With flexCust
        .Redraw = False
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
        .AllowUserResizing = flexResizeColumns
        .Editable = flexEDKbdMouse

        .ColWidth(0) = 6000
        .ColWidth(1) = 0
        .ColWidth(2) = 50
        .ColWidth(3) = 0

        .MergeCol(0) = True
        .Cell(flexcpText, 0, 0, 1, 0) = mDocTitleCust
        .Cell(flexcpFontName, 0, 0, 1, 0) = "MS Sans Serif"
'        .Cell(flexcpFontName, 0, 0, 1, 0) = "Haettenschweiler"
        .Cell(flexcpFontSize, 0, 0, 1, 0) = "10"
        .Cell(flexcpForeColor, 0, 0, 1, 0) = mBackColor
        
        ' invoice No.
        PrintTitle flexTitle, "CUSTOMER", 1, 1
        PrintTitle flexTitle, "PRE INVOICE NO.", 1, 3
        
        ' draw borders
        .Select 1, 3, 2, 3
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        .Select 1, 1, 2, 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
        
        'set column aligment
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        
        .Redraw = True
        
    End With
    
End Sub
Private Sub SetItemsGrid()
 
 Dim i%
 
 With flexItems
    
        .Redraw = False
        ' overall grid settings
        .FixedRows = 1
        .FixedCols = 0
        .Rows = 20
        .Cols = 9
        .ScrollBars = flexScrollBarBoth
        .GridLines = flexGridFlat
        .Appearance = flexFlat
        .BorderStyle = flexBorderFlat
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightNever
        .Editable = True
        .SheetBorder = vbWhite
        .BackColorBkg = vbWhite
        .GridLines = flexGridFlatVert
        .FontName = "MS Sans Serif"
        .FontSize = 8
        .AllowUserResizing = flexResizeColumns
     
        ' set column width
        .ColWidth(0) = 0
        .ColWidth(1) = 4000
        .ColWidth(2) = 0
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1200
        .ColWidth(8) = 0
        
        ' Print titles
        PrintTitle flexItems, "TRIP_NO", 0, 0
        PrintTitle flexItems, "รายการ", 0, 1
        PrintTitle flexItems, "QUOTATION", 0, 2
        PrintTitle flexItems, "น้ำหนัก", 0, 3
        PrintTitle flexItems, "หน่วย", 0, 4
        PrintTitle flexItems, "อัตราค่าบริการ", 0, 5
        PrintTitle flexItems, "หน่วย", 0, 6
        PrintTitle flexItems, "จำนวน", 0, 7
        PrintTitle flexItems, "ID", 0, 8
        .ForeColor = vbBlack
        
        ' apply common formatting
        .Select 0, 0, 0, .Cols - 1
        .CellBorder vbBlack, 1, 1, 1, 1, 1, 1
 
        .Redraw = True
    End With
End Sub
