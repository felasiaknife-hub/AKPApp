VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenifestStatusListApprove 
   BackColor       =   &H00FFFFFF&
   Caption         =   "List Menifest For Approve"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   ForeColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   9720
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   3420
      Top             =   7440
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9420
      Left            =   0
      ScaleHeight     =   9390
      ScaleWidth      =   2355
      TabIndex        =   13
      Top             =   0
      Width           =   2385
      Begin VB.TextBox txtTimeTableNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   4260
         Width           =   1995
      End
      Begin VB.ComboBox cboStatus3 
         Height          =   315
         ItemData        =   "frmMenifestStatusListApprove.frx":0000
         Left            =   180
         List            =   "frmMenifestStatusListApprove.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2040
         Width           =   2025
      End
      Begin VB.ComboBox cboStatus2 
         Height          =   315
         ItemData        =   "frmMenifestStatusListApprove.frx":0004
         Left            =   180
         List            =   "frmMenifestStatusListApprove.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1680
         Width           =   2025
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-"
         Height          =   255
         Left            =   660
         TabIndex        =   9
         Top             =   4650
         Width           =   555
      End
      Begin VB.CommandButton Command3 
         Caption         =   "+"
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   4650
         Width           =   555
      End
      Begin VB.TextBox txtKey2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Text            =   "*"
         Top             =   960
         Width           =   1995
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   5
         Top             =   2700
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.CommandButton cmdViewJobBill 
         Caption         =   "ดูข้อมูลใบ Manifest"
         Height          =   435
         Left            =   60
         TabIndex        =   11
         Top             =   5580
         Width           =   2205
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   60
         TabIndex        =   12
         Top             =   6030
         Width           =   2205
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   60
         TabIndex        =   10
         Top             =   5130
         Width           =   2205
      End
      Begin VB.ComboBox cboJobStatus 
         Height          =   315
         ItemData        =   "frmMenifestStatusListApprove.frx":0008
         Left            =   180
         List            =   "frmMenifestStatusListApprove.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   2025
      End
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   360
         Width           =   1995
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   6
         Top             =   3360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบคุม"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   24
         Top             =   4020
         Width           =   765
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   90
         X2              =   2220
         Y1              =   3900
         Y2              =   3900
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   120
         Top             =   7440
         Width           =   195
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปิดงานแล้ว"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   420
         TabIndex        =   23
         Top             =   7440
         Width           =   795
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   90
         X2              =   2220
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor, คนขับ, ทะเบียนรถ:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   720
         Width           =   1965
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   420
         TabIndex        =   21
         Top             =   7680
         Width           =   495
      End
      Begin VB.Label lbBilled 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วางบิลแล้ว"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   420
         TabIndex        =   20
         Top             =   7200
         Width           =   735
      End
      Begin VB.Label lbPrinted 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พิมพ์ Menifest แล้ว"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   420
         TabIndex        =   19
         Top             =   6960
         Width           =   1335
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "งานใหม่"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   420
         TabIndex        =   18
         Top             =   6720
         Width           =   570
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   4
         Left            =   120
         Top             =   7680
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   120
         Top             =   7200
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   120
         Top             =   6960
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   120
         Top             =   6720
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   17
         Top             =   3120
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ปฏิบัติงานตั้งแต่"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   16
         Top             =   2460
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลข Manifest, ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   14
         Top             =   120
         Width           =   1650
      End
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComctlLib.ImageList ImgSorted 
      Left            =   3900
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenifestStatusListApprove.frx":000C
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenifestStatusListApprove.frx":05A6
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VSFlex7LCtl.VSFlexGrid vsMenifestList 
      Height          =   6315
      Left            =   2400
      TabIndex        =   25
      Top             =   0
      Width           =   7695
      _cx             =   13573
      _cy             =   11139
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อรุ่น,ลาย หรือยี่ห้อ :"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   15
      Top             =   4950
      Width           =   1440
   End
End
Attribute VB_Name = "frmMenifestStatusListApprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' In General Declarations
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_HITTEST As Long = (LVM_FIRST + 18)
Private Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57)
Private Const LVHT_ONITEMICON As Long = &H2
Private Const LVHT_ONITEMLABEL As Long = &H4
Private Const LVHT_ONITEMSTATEICON As Long = &H8
Private Const LVHT_ONITEM As Long = (LVHT_ONITEMICON Or _
                                    LVHT_ONITEMLABEL Or _
                                    LVHT_ONITEMSTATEICON)

'Private Type POINTAPI
'  x As Long
'  Y As Long
'End Type

'Private Type LVHITTESTINFO
'   Pt As POINTAPI
'   flags As Long
'   iItem As Long
'   iSubItem  As Long
'End Type

Dim lX As Single, lY As Single

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim curCondition_ As String
Dim ColumnIndex As Integer

Const slComboStatus = "งานใหม่|พิมพ์ Menifest แล้ว|ปิดงานแล้ว|ยกเลิก"
Const slComboCheckDoc = "ปิดใบเมนิเฟส|ยกเลิกการปิดงาน|ยกเลิก"

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    SearchQue
End Sub

Private Sub cmdViewJobBill_Click()
    'lvResult_DblClick
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Me.ActiveControl.Container.Name = txtKey.Container.Name Then
            If Not TypeOf Me.ActiveControl Is CommandButton Then
                cmdSearch_Click
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim backDate As String
Dim tmpY As Integer, tmpM As Integer, tmpD As Integer
Dim tmpArr
Dim sqlStr As String
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    ' Hook Form
    Call WheelHook(Me.hwnd)

  ' Hook Controls to be ignored
    Dim ctl As Control
    For Each ctl In Controls
        On Error Resume Next
        Select Case True
        Case TypeOf ctl Is MSFlexGrid, TypeOf ctl Is VSFlexGrid
            Call WheelHook(ctl.hwnd)
        Case TypeOf ctl Is PictureBox
            Call WheelHook(ctl.hwnd)
        Case TypeOf ctl Is ListBox, TypeOf ctl Is ListView, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
            Call WheelHook(ctl.hwnd)
        Case Else
            '
        End Select
    Next ctl
    
    sqlStr = "select backdate=convert(varchar(10),getdate()-7,111)"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then backDate = "0"
    backDate = Trim("" & tmpRec("BackDate"))
    tmpRec.Close
    If IsDate(backDate) Then
        tmpArr = Split(backDate, "/")
        If CInt(tmpArr(0)) < 2500 Then
            dtFrom.ValueYMD = Trim("" & (CInt(tmpArr(0) + 543))) & "/" & Trim("" & tmpArr(1)) & "/" & Trim("" & tmpArr(2))
        Else
            dtFrom.ValueYMD = backDate
        End If
    End If
    Set tmpRec = Nothing
    With cboJobStatus
        .AddItem "== ทุกสถานะ =="
        .AddItem "งานใหม่"
        .AddItem "พิมพ์ Menifest แล้ว"
        .AddItem "ปิดงานแล้ว"
        .AddItem "ยกเลิก"
        .ListIndex = 0
    End With
    With cboStatus2
        .AddItem "== ทุกสถานะวางบิล =="
        .AddItem "ยังไม่ได้วางบิล"
        .AddItem "วางบิลแล้ว"
        .ListIndex = 0
    End With
    With cboStatus3
        .AddItem "== ทุกสถานะจ่าย Vendor =="
        .AddItem "ยังไม่ได้จ่าย"
        .AddItem "จ่าย Vendor แล้ว"
        .ListIndex = 0
    End With
    
    SetDefaults vsMenifestList
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error Resume Next
    With vsMenifestList
        If Me.Width - Picture1.Width > 500 Then
            .Width = Me.Width - Picture1.Width - 200
        End If
        .Height = Me.Height - .Top - 500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Stb1.Panels(1).Text = App.Title
    VisibleImage
    Call WheelUnHook(Me.hwnd)
End Sub

Private Sub methMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Pont As POINTAPI
    Dim ret As BTN_STYLE

    If Button = vbRightButton Then
        hMenu = CreatePopupMenu()
        AppendMenu hMenu, MF_STRING, 1, "Auto Calculate Manifest"
        AppendMenu hMenu, MF_SEPARATOR, -1, "SEP1"
        AppendMenu hMenu, MF_STRING, 2, "แสดงเฉพาะรายการที่มียอดเป็น 0"
'        AppendMenu hMenu, MF_SEPARATOR, -1, "SEP2"
'        AppendMenu hMenu, MF_STRING, 3, "ไม่อนุมัติจ่าย"
'        If LCase(CurrentUser.LoginName) = "admin" Then AppendMenu hMenu, MF_STRING, 4, "อนุมัติจ่าย"
        GetCursorPos Pont
        ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD, _
                                Pont.x, Pont.y, Me.hwnd, ByVal 0&)
        DestroyMenu hMenu

'        If ret = 1 Then
'            MNU_AutoCalcMF
'        ElseIf ret = 2 Then
'            MNU_Show_Only_MF0
'        ElseIf ret = 3 Then
'            MNU_NotApprove_Click
'        ElseIf ret = 4 Then
'            MNU_SetApprove_Click
'        End If
    End If
End Sub

Private Sub SearchQue(Optional ByVal ShowActualQty0 As Boolean = False)
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim tmpList As ListItem
Dim x As Long, y As Long
Dim sqlStr As String, strCond As String, StatusDesc As String, UpdateStatus As String, InvStatus As String
Dim tmpArr
Dim curColor
Dim subStr As String
    SetDefaults vsMenifestList
    Screen.MousePointer = 11
    sqlStr = "": strCond = ""
    sqlStr = "SELECT MenifestID,DocNo,MenifestNo,CustomerName,SubTypeDesc,TruckTypeDesc,TransportName1,TransportName2,WorkDate,isClosed,isCanceled,isPrinted,DriverName,JobNo,DocJobRecord,DocControl,DocSiteWeight,DocCustWeight,DocCustControl,BillPayHeaderID,BillPayNo,BillingHeaderID,BillingNo,isApproved, CASE WHEN ISNULL(isCancelHead, 'Y') = 'Y' THEN 'N' ELSE CASE WHEN ISNULL(BillingNo, '') = '' THEN ISNULL(isInv, 'N') ELSE 'Y' END END AS isInv, isCancelHead AS InvIsCancel From vw_MenifestQue " & vbCrLf
    If Trim(txtTimeTableNo.Text) = "" Then
        strCond = " WHERE (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "') And (isClosed='Y') And (isCanceled<>'Y' Or isCanceled is null) " & vbCrLf
        If Trim(txtKey.Text) <> "" And Replace(Trim(txtKey.Text), "*", "") <> "" Then
            strCond = strCond & " And (MenifestNo Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or DocNo Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or CustomerName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')" & vbCrLf
        End If
        If Trim(txtKey2.Text) <> "" And Replace(Trim(txtKey2.Text), "*", "") <> "" Then
            strCond = strCond & " And (TransportName1 Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or TransportName2 Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or DriverName Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or CarRegisID Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%')" & vbCrLf
        End If
        Select Case cboJobStatus.ListIndex
            Case 1 'งานใหม่
                subStr = " (ISNULL(isPrinted, 'N') <> 'Y') "
            Case 2 'พิมพ์ ใบ Menifest แล้ว
                subStr = " (ISNULL(isPrinted, 'N') = 'Y') "
            Case 3 'ปิดงานแล้ว
                subStr = " (isClosed = 'Y') "
            Case 4 'ยกเลิก
                subStr = " (isCanceled = 'Y') "
        End Select
        If cboJobStatus.ListIndex > 0 Then
            strCond = strCond & " AND " & subStr
        End If
        Select Case cboStatus2.ListIndex
            Case 1 'ยังไม่ได้วางบิล
'                subStr = " (BillingHeaderID is null) "
                subStr = "(CASE WHEN ISNULL(isCancelHead, 'Y') = 'Y' THEN 'N' ELSE CASE WHEN ISNULL(BillingNo, '') = '' THEN ISNULL(isInv, 'N') ELSE 'Y' END END <> 'Y')" & vbCrLf
            Case 2 'วางบิลแล้ว
'                subStr = " (BillingHeaderID is not null) "
                subStr = " (isInv = 'Y')  AND (ISNULL(BillingHeaderID, '') <> '') "
        End Select
        If cboStatus2.ListIndex > 0 Then
            strCond = strCond & " And " & subStr
        End If
        Select Case cboStatus3.ListIndex
            Case 1 'ยังไม่ได้จ่าย
                subStr = " (BillPayHeaderID is null) "
            Case 2 'จ่ายแล้ว
                subStr = " (BillPayHeaderID is not null) "
        End Select
        If cboStatus3.ListIndex > 0 Then
            strCond = strCond & " And " & subStr
        End If
    Else
        strCond = " WHERE (TimeTableNo = '" & txtTimeTableNo.Text & "') "
    End If
    If ShowActualQty0 = False Then
'        strCond = strCond & " AND (ActualQty > 0)"
    Else
        strCond = strCond & " AND (ActualQty > 0) AND (IncollectFee = 0)"
    End If
    curCondition_ = strCond
    
    sqlStr = sqlStr & strCond
    sqlStr = SelectWithCompanyID(sqlStr, False)
    sqlStr = sqlStr & " ORDER BY WorkDate, MenifestNo, DocNo "
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ  จำนวน 0 เที่ยว"
        Exit Sub
    End If
    If tmpRec.RecordCount = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ  จำนวน 0 เที่ยว"
        Exit Sub
    End If
    
    SetDefaults vsMenifestList
    Dim TotalRecord As Long
    TotalRecord = 0
    If tmpRec.EOF = False Then
        TotalRecord = tmpRec.RecordCount
        Do While Not tmpRec.EOF
            StatusDesc = "งานใหม่"
            UpdateStatus = "ปิดใบเมนิเฟส"
            InvStatus = "ยังไม่ได้วางบิล"
            With vsMenifestList
                If tmpRec!IsCanceled = "Y" Then
                    curColor = lbCancel.ForeColor
                    StatusDesc = lbCancel.Caption
                    UpdateStatus = lbCancel.Caption
                ElseIf tmpRec!isClosed = "Y" Then
                    curColor = lbClosed.ForeColor
                    StatusDesc = lbClosed.Caption
                    UpdateStatus = "ปิดใบเมนิเฟส"
                ElseIf tmpRec!isPrinted = "Y" Then
                    curColor = lbPrinted.ForeColor
                    StatusDesc = lbPrinted.Caption
                ElseIf tmpRec!isPrinted <> "Y" Then
                    curColor = lbOpen.ForeColor
                    StatusDesc = lbOpen.Caption
                Else
                    curColor = vbBlack
                End If
                
                If tmpRec!isInv = "Y" And tmpRec!BillingNo <> "" Then
                    InvStatus = lbBilled.Caption
                End If
                
                .AddItem tmpRec.AbsolutePosition & vbTab & tmpRec!menifestID & vbTab & tmpRec!docNo & vbTab & tmpRec!MenifestNo & vbTab & tmpRec!CustomerName & vbTab & tmpRec!WorkDate & vbTab & tmpRec!JobNo & vbTab & StatusDesc & vbTab & InvStatus & vbTab & IIf(tmpRec!isApproved = "N", 0, 1)
                SetRowStyle StatusDesc, .Rows - 1
            End With
            DoEvents
            tmpRec.MoveNext
        Loop
    End If
    Screen.MousePointer = 0
    frmMain.Stb1.Panels(1).Text = "พบข้อมูล " & Format(TotalRecord, "#,##0") & " รายการ"
    tmpRec.Close
    Set tmpRec = Nothing
    Exit Sub
ErrD:
    Screen.MousePointer = 0
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub SetRowStyle(ByVal Status As String, ByVal Row As Long)
Dim curColor, i#
    Select Case Status
        Case "งานใหม่"
            curColor = Shape1(0).FillColor
        Case "พิมพ์ Menifest แล้ว"
            curColor = Shape1(1).FillColor
        Case "วางบิลแล้ว"
            curColor = Shape1(2).FillColor
        Case "ปิดงานแล้ว"
            curColor = Shape1(3).FillColor
        Case "ยกเลิก"
            curColor = Shape1(4).FillColor
        Case Else
            curColor = vbBlack
    End Select
    For i = 0 To vsMenifestList.Cols - 1
        vsMenifestList.Row = Row
        vsMenifestList.Col = i
        vsMenifestList.CellForeColor = curColor
    Next
'    vsMenifestList.TextMatrix(Row, vsMenifestList.Cols - 1) = Status
End Sub

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  Dim bHandled As Boolean
  Dim bOver As Boolean
  
  For Each ctl In Controls
    ' Is the mouse over the control
    On Error Resume Next
    bOver = (ctl.Visible And IsOver(ctl.hwnd, Xpos, Ypos))
    On Error GoTo 0
    
    If bOver Then
      ' If so, respond accordingly
      bHandled = True
      Select Case True
      
        Case TypeOf ctl Is MSFlexGrid
          FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
          
        Case TypeOf ctl Is VSFlexGrid
          VSFlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
          
        Case TypeOf ctl Is PictureBox
          PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
          
        Case TypeOf ctl Is ListBox, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
          ' These controls already handle the mousewheel themselves, so allow them to:
          If ctl.Enabled Then ctl.SetFocus
          
        Case Else
          bHandled = False

      End Select
      If bHandled Then Exit Sub
    End If
    bOver = False
  Next ctl
  
  ' Scroll was not handled by any controls, so treat as a general message send to the form
  Me.Caption = "Form Scroll " & IIf(Rotation < 0, "Down", "Up")
End Sub

Private Sub SetDefaults(fa As VSFlexGrid)
    DoEvents
    With fa
        .Clear
        .Rows = 1
        .Cols = 10
        .FixedRows = 1
        .FixedCols = 0
        .BindToArray Null
        .ScrollTrack = False
        .ExplorerBar = flexExNone
        .AutoSearch = flexSearchNone
        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeNone
        .SelectionMode = flexSelectionFree
        .OutlineBar = flexOutlineBarNone
        .OLEDragMode = flexOLEDragManual
        .OLEDropMode = flexOLEDropNone
        .ScrollTips = False
        .ToolTipText = ""
        
        .ColWidth(0) = 500
        .ColWidth(1) = 0
        .ColWidth(2) = 1700
        .ColWidth(3) = 0
        .ColWidth(4) = 3000
        .ColWidth(5) = 1200
        .ColWidth(6) = 1400
        .ColWidth(7) = 1700
        .ColWidth(8) = 1700
        .ColWidth(9) = 1100
        
        'MenifestID, DocNo, MenifestNo, CustomerName, WorkDate, isClosed, isCanceled,isPrinted,DriverName,JobNo,DocJobRecord,DocControl,DocSiteWeight,DocCustWeight,DocCustControl,BillPayHeaderID,BillPayNo,BillingHeaderID,BillingNo,isDocCompleted, CASE WHEN ISNULL(isCancelHead, 'Y') = 'Y' THEN 'N' ELSE CASE WHEN ISNULL(BillingNo, '') = '' THEN ISNULL(isInv, 'N') ELSE 'Y' END END AS isInv, isCancelHead AS InvIsCancel
        
        .TextMatrix(0, 0) = "#"
        .TextMatrix(0, 1) = "ManifestID"
        .TextMatrix(0, 2) = "DocNo"
        .TextMatrix(0, 3) = "ManifestNo"
        .TextMatrix(0, 4) = "Customer"
        .TextMatrix(0, 5) = "WorkDate"
        .TextMatrix(0, 6) = "JobNo"
        .TextMatrix(0, 7) = "สถานะเมนิเฟส"
        .TextMatrix(0, 8) = "สถานะวางบิล"
        .TextMatrix(0, 9) = "Approve"
        
        .ColDataType(0) = flexDTString
        .ColDataType(1) = flexDTString
        .ColDataType(2) = flexDTString
        .ColDataType(3) = flexDTString
        .ColDataType(4) = flexDTString
        .ColDataType(5) = flexDTString
        .ColDataType(6) = flexDTString
        .ColDataType(7) = flexDTString
        .ColDataType(8) = flexDTString
        .ColDataType(9) = flexDTBoolean
        
        .ColComboList(7) = slComboStatus
'        .ColComboList(8) = slComboCheckDoc
        
        .ColFormat(5) = "##/##/####"
        .ColEditMask(5) = "00/00/0000"
        
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(9) = flexAlignCenterCenter
        
        ' set the properties we want
        .AllowUserResizing = flexResizeBoth
        .ExtendLastCol = True
        .SelectionMode = flexSelectionFree
    End With
End Sub

Private Sub vsMenifestList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim tmpS As New BWGSearchManager
Dim myMenifest_ As BWGMenifest
Dim tmpCol As Collection
Dim UpdateLog As String
    If Row <= 0 Then Exit Sub
    If Trim(vsMenifestList.TextMatrix(Row, 1)) = "" Then Exit Sub
    Set tmpCol = tmpS.MenifestSearch("MenifestID='" & vsMenifestList.TextMatrix(Row, 1) & "'")
    If tmpCol.Count > 0 Then
        Set myMenifest_ = tmpCol(1)
        Select Case Col
'            Case 8
'                If vsMenifestList.TextMatrix(Row, Col) = "ปิดงานแล้ว" Then
''                    myMenifest_.isClosed = True
'                    SetRowStyle lbClosed.Caption, Row
'                ElseIf vsMenifestList.TextMatrix(Row, Col) = "ยกเลิกการปิดงาน" Then
''                    myMenifest_.isClosed = False
'                    SetRowStyle lbPrinted.Caption, Row
'                ElseIf vsMenifestList.TextMatrix(Row, Col) = "ยกเลิกเมนิเฟส" Then
''                    myMenifest_.IsCanceled = True
'                    SetRowStyle lbCancel.Caption, Row
'                End If
            Case 9
                If vsMenifestList.TextMatrix(Row, Col) = "-1" Then vsMenifestList.TextMatrix(Row, Col) = "1"
                If vsMenifestList.TextMatrix(Row, Col) = "1" Then vsMenifestList.TextMatrix(Row, Col) = "1"
                If vsMenifestList.TextMatrix(Row, Col) = "0" Then vsMenifestList.TextMatrix(Row, Col) = "0"
                If vsMenifestList.TextMatrix(Row, Col) = "1" Then
                    myMenifest_.isApproved = True
                Else
                    myMenifest_.isApproved = False
                End If
                UpdateLog = "Menifest is Approve = " & CBool(vsMenifestList.TextMatrix(Row, Col))
            End Select
        UpdateLog = UpdateLog & " " & myMenifest_.ApproveMenifest
        frmMain.Stb1.Panels(1).Text = " Update " & myMenifest_.docNo & " (" & UpdateLog & ")"
    End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
End Sub

Private Sub vsMenifestList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
        Case 0, 1, 2, 3, 4, 5, 6.7, 8
            Cancel = True
            Exit Sub
    End Select
End Sub

Private Sub vsMenifestList_DblClick()
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim r&, c&
    r = vsMenifestList.MouseRow
    c = vsMenifestList.MouseCol
    If r <= 0 Then Exit Sub
    If Trim(vsMenifestList.TextMatrix(r, 1)) = "" Then Exit Sub
    Set tmpCol = tmpS.MenifestSearch("MenifestID='" & vsMenifestList.TextMatrix(r, 1) & "'")
    If tmpCol.Count > 0 Then
        Load frmMenifest
        frmMenifest.SetDataMenifest tmpCol(1)
        frmMenifest.SSTab1.TabVisible(2) = True
        frmMenifest.Show 1
    End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
End Sub
