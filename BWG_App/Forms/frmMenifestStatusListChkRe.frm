VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenifestStatusListChkRe 
   BackColor       =   &H00FFFFFF&
   Caption         =   "List Invoice Check Customer Reciev"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12360
   ForeColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   12360
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
      TabIndex        =   7
      Top             =   0
      Width           =   2385
      Begin VB.ComboBox cboCompany 
         Height          =   315
         ItemData        =   "frmMenifestStatusListChkRe.frx":0000
         Left            =   180
         List            =   "frmMenifestStatusListChkRe.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1080
         Width           =   2055
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   3
         Top             =   2460
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   60
         TabIndex        =   6
         Top             =   4350
         Width           =   2205
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   60
         TabIndex        =   5
         Top             =   3810
         Width           =   2205
      End
      Begin VB.ComboBox cboJobStatus 
         Height          =   315
         ItemData        =   "frmMenifestStatusListChkRe.frx":0004
         Left            =   180
         List            =   "frmMenifestStatusListChkRe.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1740
         Width           =   2025
      End
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   420
         Width           =   1995
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   4
         Top             =   3120
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label LbIsFile 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ส่งฟ้อง"
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   6600
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000C0C0&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   120
         Top             =   6600
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานะ"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   1500
         Width           =   465
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   90
         X2              =   2220
         Y1              =   3660
         Y2              =   3660
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   390
         TabIndex        =   14
         Top             =   6210
         Width           =   480
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้ารับวางบิลแล้ว"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   390
         TabIndex        =   13
         Top             =   5850
         Width           =   1305
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้ายังไม่รับวางบิล"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   390
         TabIndex        =   12
         Top             =   5490
         Width           =   1380
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   120
         Top             =   6180
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   120
         Top             =   5820
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   120
         Top             =   5460
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   11
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ปฏิบัติงานตั้งแต่"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   10
         Top             =   2220
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลข Invoice, ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   8
         Top             =   120
         Width           =   1575
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
            Picture         =   "frmMenifestStatusListChkRe.frx":0008
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenifestStatusListChkRe.frx":05A2
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VSFlex7LCtl.VSFlexGrid vsMenifestList 
      Height          =   6315
      Left            =   2400
      TabIndex        =   15
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
      TabIndex        =   9
      Top             =   4950
      Width           =   1440
   End
End
Attribute VB_Name = "frmMenifestStatusListChkRe"
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



Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    SearchQue
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
    
    PopulateCompany cboCompany, cboIndex
    
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
        .AddItem "ลูกค้ายังไม่รับวางบิล"
        .AddItem "ลูกค้ารับวางบิล"
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
'     UnHook Form
    Call WheelUnHook(Me.hwnd)
End Sub

Private Sub methMouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
                                Pont.x, Pont.Y, Me.hwnd, ByVal 0&)
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
Dim x As Long, Y As Long
Dim sqlStr As String, strCond As String, StatusDesc As String
Dim tmpArr
Dim curColor
Dim subStr As String
    SetDefaults vsMenifestList
    Screen.MousePointer = 11
    sqlStr = "SELECT BillingHeaderID, CompanyCode,  CustomerCode, BillName, BillingNo, IssuedDate,PayDueDate, ISNULL(isCustAccepted, 'N') AS isCustAccepted, isCancel,ISNull(isFiled, 'N') as isFiled,DatePast,ISNull(IsEMS, 'N') as IsEMS,ISNull(IsMan, 'N') as IsMan,ManName,preinvno" & vbCrLf
    sqlStr = sqlStr & "FROM vw_BillHeaderSearch " & vbCrLf
    strCond = "WHERE (IssuedDate>='" & dtFrom.ValueYMD & "' And IssuedDate<='" & dtTo.ValueYMD & "') AND (isCancel <> 'Y' OR isCancel IS NULL) " & vbCrLf
    
    If Trim(txtKey.Text) <> "" And Replace(Trim(txtKey.Text), "*", "") <> "" Then
        strCond = strCond & " AND (BillingNo LIKE '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' OR CompanyName LIKE '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or BillName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')" & vbCrLf
    End If
    
    If cboCompany.ListCount > 1 Then
        If cboCompany.ListIndex > 0 Then
            strCond = strCond & "AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' " & vbCrLf
        End If
    End If
    
    Select Case cboJobStatus.ListIndex
        Case 1 'ยังไม่ได้รับวางบิล
            subStr = " (ISNULL(isCustAccepted, 'N') = 'N') "
        Case 2 'รับวางบิลแล้ว
            subStr = " (ISNULL(isCustAccepted, 'N') = 'Y') "
    End Select
    
    If cboJobStatus.ListIndex > 0 Then
        strCond = strCond & " And " & subStr
    End If
    
    curCondition_ = strCond
    
    sqlStr = sqlStr & strCond
    '*** ใส่ CompanyID ให้อัตโนมัติ ตามสิทธิการใช้งานของแต่ละบริษัท
    'sqlStr = SelectWithCompanyID(sqlStr, False)
    '***********
    sqlStr = sqlStr & vbCrLf & "ORDER BY CompanyCode, IssuedDate, BillingNo"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ"
        Exit Sub
    End If
    If tmpRec.RecordCount = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ"
        Exit Sub
    End If
    SetDefaults vsMenifestList
    Dim TotalRecord As Long
    Dim BillNo As String
    TotalRecord = 0
    If tmpRec.EOF = False Then
        TotalRecord = tmpRec.RecordCount
        Do While Not tmpRec.EOF
            With vsMenifestList
                If tmpRec!isCancel = "Y" Then
                    curColor = lbCancel.ForeColor
                    StatusDesc = lbCancel.Caption
                ElseIf Trim("" & tmpRec!isCustAccepted) = "N" Then
                    curColor = lbOpen.ForeColor
                    StatusDesc = lbOpen.Caption
                ElseIf tmpRec!isCustAccepted = "Y" Then
                    curColor = lbApproved.ForeColor
                    StatusDesc = lbApproved.Caption
                ElseIf Trim("" & tmpRec!isCustAccepted) = "" Then
                    curColor = lbOpen.ForeColor
                    StatusDesc = lbOpen.Caption
                Else
                    curColor = vbBlack
                    StatusDesc = "งานใหม่"
                End If
                If Trim("" & tmpRec!isFiled) = "Y" Then
                    curColor = LbIsFile.ForeColor
                End If
                If tmpRec!BillingNo <> "" Then
                    BillNo = tmpRec!BillingNo
                Else
                    BillNo = tmpRec!PreInvNo
                End If
                .AddItem tmpRec.AbsolutePosition & vbTab & tmpRec!BillingHeaderID & vbTab & tmpRec!CompanyCode & vbTab & tmpRec!CustomerCode & vbTab & tmpRec!BillName & vbTab & tmpRec!BillingNo & vbTab & Format(tmpRec!IssuedDate, "dd/MM/yyyy") & vbTab & Format(tmpRec!PayDueDate, "dd/MM/yyyy") & vbTab & IIf(Trim(tmpRec!isCustAccepted) = "", 0, IIf(tmpRec!isCustAccepted = "N", 0, 1)) & vbTab & StatusDesc & vbTab & IIf(Trim(tmpRec!isFiled) = "Y", 1, 0) & vbTab & tmpRec!DatePast & vbTab & IIf(Trim(tmpRec!IsEMS) = "Y", 1, 0) & vbTab & IIf(Trim(tmpRec!IsMan) = "Y", 1, 0) & vbTab & "                        " & tmpRec!ManName & vbTab
'                SetColor vsMenifestList, .Rows - 1, vbBlue
                    SetRowStyle StatusDesc, curColor, .Rows - 1
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
Private Sub SetColor(Flexgrid As MSFlexGrid, rowNo As Integer, bkColor As ColorConstants)
With Flexgrid
    If rowNo >= .Rows Then
    Exit Sub
    End If
    .FillStyle = flexFillRepeat
    .col = .FixedCols
    .RowSel = rowNo
    .ColSel = .col - 1
    .CellBackColor = bkColor
End With
End Sub

Private Sub SetRowStyle(ByVal Status As String, ByVal Colors As Long, ByVal row As Long)
Dim curColor, i#
    curColor = Colors
    For i = 0 To vsMenifestList.Cols - 1
        vsMenifestList.row = row
        vsMenifestList.col = i
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
        .Cols = 15
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
        .ColWidth(2) = 800
        .ColWidth(3) = 1000
        .ColWidth(4) = 3000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 1500
        .ColWidth(10) = 500
        .ColWidth(11) = 1000
        .ColWidth(12) = 500
        .ColWidth(13) = 800
        .ColWidth(14) = 1000
        
        .TextMatrix(0, 0) = "#"
        .TextMatrix(0, 1) = "BillingHeaderID"
        .TextMatrix(0, 2) = "Company"
        .TextMatrix(0, 3) = "CustCode"
        .TextMatrix(0, 4) = "Customer"
        .TextMatrix(0, 5) = "Inv No."
        .TextMatrix(0, 6) = "IssuedDate"
        .TextMatrix(0, 7) = "PayDueDate"
        .TextMatrix(0, 8) = "รับวางบิล"
        .TextMatrix(0, 9) = "สถานะ"
        .TextMatrix(0, 10) = "ส่งฟ้อง"
        .TextMatrix(0, 11) = "วันวางบิล"
        .TextMatrix(0, 12) = "EMS"
        .TextMatrix(0, 13) = "วางบิลเอง"
        .TextMatrix(0, 14) = "พนักงาน"
        
        .ColDataType(0) = flexDTString
        .ColDataType(1) = flexDTString
        .ColDataType(2) = flexDTString
        .ColDataType(3) = flexDTString
        .ColDataType(4) = flexDTString
        .ColDataType(5) = flexDTString
        .ColDataType(6) = flexDTString
        .ColDataType(7) = flexDTString
        .ColDataType(8) = flexDTBoolean
        .ColDataType(9) = flexDTString
        .ColDataType(10) = flexDTBoolean
        .ColDataType(11) = flexDTString
        .ColDataType(12) = flexDTBoolean
        .ColDataType(13) = flexDTBoolean
        .ColDataType(14) = flexDTString
        
                
        .ColFormat(6) = "##/##/####"
        .ColFormat(7) = "##/##/####"
        .ColEditMask(6) = "00/00/0000"
        .ColEditMask(7) = "00/00/0000"
        
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignCenterCenter
        .ColAlignment(10) = flexAlignCenterCenter
        .ColAlignment(11) = flexAlignCenterCenter
        .ColAlignment(12) = flexAlignCenterCenter
        .ColAlignment(13) = flexAlignCenterCenter
        .ColAlignment(14) = flexAlignCenterCenter

        ' set the properties we want
        .AllowUserResizing = flexResizeBoth
        .ExtendLastCol = True
        .SelectionMode = flexSelectionFree
    End With
End Sub

Private Sub vsMenifestList_AfterEdit(ByVal row As Long, ByVal col As Long)
Dim tmpS As New BWGSearchManager
Dim myInvoice As BWGBillingHeader
Dim tmpCol As Collection
Dim UpdateLog As String, StatusDesc As String, curColor As Long
    If row <= 0 Then Exit Sub
    If Trim(vsMenifestList.TextMatrix(row, 1)) = "" Then Exit Sub
    Set tmpCol = tmpS.BillingHeaderSearch("BillingHeaderID='" & vsMenifestList.TextMatrix(row, 1) & "'")
    If tmpCol.Count > 0 Then
        Set myInvoice = tmpCol(1)
        If myInvoice.isCancel = True Then Exit Sub
        Select Case col
            Case 7, 8
                myInvoice.PayDueDate = Format(vsMenifestList.TextMatrix(row, 7), "yyyy/MM/dd")
                UpdateLog = "PayDueDate = " & Format(vsMenifestList.TextMatrix(row, 7), "dd/MM/yyyy")
                If vsMenifestList.TextMatrix(row, 8) = "-1" Then vsMenifestList.TextMatrix(row, 8) = "1"
                If vsMenifestList.TextMatrix(row, 8) = "1" Then vsMenifestList.TextMatrix(row, 8) = "1"
                If vsMenifestList.TextMatrix(row, 8) = "0" Then vsMenifestList.TextMatrix(row, 8) = "0"
                If vsMenifestList.TextMatrix(row, 8) = "1" Then
                    myInvoice.isCustAccepted = True
                    StatusDesc = lbApproved.Caption
                    curColor = lbApproved.ForeColor
                Else
                    myInvoice.isCustAccepted = False
                    StatusDesc = lbOpen.Caption
                    curColor = lbOpen.ForeColor
                End If
                UpdateLog = UpdateLog & " isCustAccepted = " & myInvoice.isCustAccepted
                SetRowStyle StatusDesc, curColor, row
                UpdateLog = UpdateLog & " " & myInvoice.UpdateReceive
            End Select
             Select Case col
                Case 10
                    If vsMenifestList.TextMatrix(row, 10) = True Then
                    DBConnExc "Update tbBillingHeader Set isFiled = 'Y'  WHERE BillingHeaderID='" & vsMenifestList.TextMatrix(row, 1) & "'"
                    curColor = LbIsFile.ForeColor
                    vsMenifestList.TextMatrix(row, 10) = True
                Else
                    DBConnExc "Update tbBillingHeader Set isFiled = 'N'  WHERE BillingHeaderID='" & vsMenifestList.TextMatrix(row, 1) & "'"
                    StatusDesc = lbOpen.Caption
                    curColor = lbOpen.ForeColor
                     vsMenifestList.TextMatrix(row, 10) = False
                    SetRowStyle StatusDesc, curColor, row
                End If
                 Case 11
                    DBConnExc "Update tbBillingHeader Set DatePast = '" & vsMenifestList.TextMatrix(row, 11) & "'  WHERE BillingHeaderID='" & vsMenifestList.TextMatrix(row, 1) & "'"
                    StatusDesc = lbOpen.Caption
                    curColor = lbOpen.ForeColor
                    SetRowStyle StatusDesc, curColor, row
                 Case 12
                    If vsMenifestList.TextMatrix(row, 12) = True Then
                        DBConnExc "Update tbBillingHeader Set IsEMS = 'Y'  WHERE BillingHeaderID='" & vsMenifestList.TextMatrix(row, 1) & "'"
                        vsMenifestList.TextMatrix(row, 12) = True
                    Else
                        DBConnExc "Update tbBillingHeader Set IsEMS = 'N'  WHERE BillingHeaderID='" & vsMenifestList.TextMatrix(row, 1) & "'"
                         vsMenifestList.TextMatrix(row, 12) = False
                    End If
                 Case 13
                    If vsMenifestList.TextMatrix(row, 13) = True Then
                        DBConnExc "Update tbBillingHeader Set IsMan = 'Y'  WHERE BillingHeaderID='" & vsMenifestList.TextMatrix(row, 1) & "'"
                        vsMenifestList.TextMatrix(row, 13) = True
                    Else
                        DBConnExc "Update tbBillingHeader Set IsMan = 'N'  WHERE BillingHeaderID='" & vsMenifestList.TextMatrix(row, 1) & "'"
                         vsMenifestList.TextMatrix(row, 13) = False
                    End If
                 Case 14
                    DBConnExc "Update tbBillingHeader Set ManName = '" & Trim(vsMenifestList.TextMatrix(row, 14)) & "'  WHERE BillingHeaderID='" & vsMenifestList.TextMatrix(row, 1) & "'"
            End Select
        frmMain.Stb1.Panels(1).Text = " Update " & myInvoice.BillingNo & " (" & UpdateLog & ")"
    End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
End Sub

Private Sub vsMenifestList_BeforeEdit(ByVal row As Long, ByVal col As Long, Cancel As Boolean)
    Select Case col
        Case 0, 1, 2, 3, 4, 5, 6, 9
            Cancel = True
            Exit Sub
'        Case 7
'            If IsNumeric(vsMenifestList.TextMatrix(Row, Col)) = False Then
'                Cancel = True
'            End If
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
    Set tmpCol = tmpS.BillingHeaderSearch("BillingHeaderID='" & vsMenifestList.TextMatrix(r, 1) & "'")
    If tmpCol.Count > 0 Then
        Load frmInvoice
        frmInvoice.SetDataBillHeader tmpCol(1)
        frmInvoice.Show 1
    End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
End Sub
