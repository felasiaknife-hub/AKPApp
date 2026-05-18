VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmQuotationList 
   Caption         =   "Quotation List"
   ClientHeight    =   9795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   11055
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport rpt1 
      Left            =   4980
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9795
      Left            =   0
      ScaleHeight     =   9765
      ScaleWidth      =   2625
      TabIndex        =   10
      Top             =   0
      Width           =   2655
      Begin VB.ComboBox cboCompany 
         Height          =   315
         ItemData        =   "frmQuotationList.frx":0000
         Left            =   180
         List            =   "frmQuotationList.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   2325
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmQuotationList.frx":0004
         Left            =   180
         List            =   "frmQuotationList.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2160
         Width           =   2325
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   180
         TabIndex        =   6
         Top             =   4080
         Width           =   2295
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   180
         TabIndex        =   8
         Top             =   5160
         Width           =   2295
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "พิมพ์รายงานใบเสนอราคา"
         Height          =   435
         Left            =   180
         TabIndex        =   7
         Top             =   4560
         Width           =   2295
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
      Begin DateCtl.Date dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   4
         Top             =   2790
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin DateCtl.Date dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   5
         Top             =   3450
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัทที่ออกใบเสนอราคา"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   20
         Top             =   1320
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "สถานะใบเสนอราคา :"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   1920
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสหรือชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   18
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ใบเสนอราคา :"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   2550
         Width           =   1320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   16
         Top             =   3210
         Width           =   495
      End
      Begin VB.Shape Shape1 
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   450
         TabIndex        =   15
         Top             =   8220
         Width           =   570
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อนุมัติแล้ว"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   450
         TabIndex        =   14
         Top             =   8520
         Width           =   720
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปิดใบเสนอราคา"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   450
         TabIndex        =   13
         Top             =   8820
         Width           =   1125
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   450
         TabIndex        =   12
         Top             =   9120
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย :"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Line Line2 
         X1              =   180
         X2              =   2490
         Y1              =   3960
         Y2              =   3960
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2640
      TabIndex        =   9
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImgSorted"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   20
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "QuotationNo"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CustCode"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Customer"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Truck"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Transport Fee"
         Object.Width           =   2618
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Transport Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Rate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Labor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Min Weight"
         Object.Width           =   1984
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Min Price"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Max Weight"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Max Price"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "WasteNo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "WasteName"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "isIncTrans"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Type"
         Object.Width           =   1940
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgSorted 
      Left            =   6720
      Top             =   7920
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
            Picture         =   "frmQuotationList.frx":0008
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmQuotationList.frx":05A2
            Key             =   "down"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmQuotationList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaleStaff_ As BWGCompanyStaff
Dim curCustomer_ As BWGCustomer
Dim colSearchResult_ As Collection

Private Function SQLString() As String
    Dim strCond As String
    Dim criStr As String
    criStr = Replace(Replace(Trim(txtCustomer.Text), "*", "%"), "'", "")
'    strCond = "(QuotationNo Like '%" & criStr & "%' Or CustomerCode Like '%" & criStr & "%' Or CompanyName Like '%" & criStr & "%')"
    strCond = "(IssueDate BETWEEN '" & dtFrom.ValueYMD & "' AND '" & dtTo.ValueYMD & "')"
    If Not curCustomer_ Is Nothing Then strCond = strCond & vbCrLf & "AND CustomerID = '" & curCustomer_.ID & "'"
    If Not SaleStaff_ Is Nothing Then strCond = strCond & vbCrLf & "AND (UnderSaleStaffID = '" & SaleStaff_.ID & "')"
    
    If cboStatus.ListIndex = 1 Then
        strCond = strCond & vbCrLf & "AND ISNULL(isApproved, 'N') + ISNULL(isClosed, 'N') + ISNULL(isCancel, 'N') = 'NNN'"
    ElseIf cboStatus.ListIndex = 2 Then
        strCond = strCond & vbCrLf & "AND ISNULL(isApproved, 'N') = 'Y'"
    ElseIf cboStatus.ListIndex = 3 Then
        strCond = strCond & vbCrLf & "AND ISNULL(isClosed, 'N') = 'Y'"
    ElseIf cboStatus.ListIndex = 4 Then
        strCond = strCond & vbCrLf & "AND ISNULL(isCancel, 'N') = 'Y'"
    End If
    
    If cboCompany.ListCount > 1 Then
        If cboCompany.ListIndex > 0 Then
            strCond = strCond & "AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' " & vbCrLf
        End If
    End If
    SQLString = strCond
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Dim tmpSql As String
    Dim selResult
    Dim tmpStr As String
    With rpt1
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .WindowTitle = "Quotation List"
        .ReportFileName = App.Path & "\Reports\QuotationLst.rpt"
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .Formulas(0) = "RangDate = " & Chr(34) & dtFrom.ValueDMY & " To " & dtTo.ValueDMY & Chr(34)
        tmpSql = "SELECT * FROM Com_vwQuotationSearch"
        tmpSql = tmpSql & vbCrLf & "WHERE "
        tmpSql = tmpSql & Replace(SQLString, "UnderSaleStaffID", "SaleStaffID")
        tmpSql = tmpSql & vbCrLf & "ORDER BY IssueDate, QuotationNo, RowPos"
        .SQLQuery = tmpSql
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrD
    Dim criStr As String
    Dim strCond As String
    Dim tmpRec As New ADODB.Recordset
    Dim x%, result%, Y%, i%
    Dim tmpList As ListItem
    Dim tmpS As BWGSearchManager
    Dim tmpItem As BWGQuotationHeader
    Dim tmpQuoDet As BWGQuotationDetail
    Dim tmpColDet As Collection
    
    Screen.MousePointer = 11
    Set colSearchResult_ = Nothing
    Set colSearchResult_ = New Collection
    lvResult.ListItems.Clear
    Set tmpS = New BWGSearchManager
    strCond = SQLString
    
'    strCond = SelectWithCompanyID(strCond, True)
'    Set colSearchResult_ = tmpS.QuotationDetailSearch(strCond)
'    Set tmpS = Nothing
'    Set tmpItem = colSearchResult_(x)
'    Set tmpColDet = tmpRec!QuotationDetails
    sqlStr = "SELECT * FROM vw_QuotationDetailSearch" & vbCrLf
    sqlStr = sqlStr & "WHERE " & strCond & vbCrLf
    sqlStr = sqlStr & "ORDER BY  IssueDate, QuotationNo, RowPos"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    
    If tmpRec.RecordCount = 0 Then Screen.MousePointer = 0: Exit Sub
    frmMain.Stb1.Panels(1).Text = "พบข้อมูลทั้หมด " & Format(tmpRec.RecordCount, "#,##0") & " รายการ"
    Screen.MousePointer = 0
    DoEvents
    Screen.MousePointer = 11
    Do While Not tmpRec.EOF
        Dim ForeColor
        Dim sqlSelTruck As String
        Dim tmpSel As New ADODB.Recordset
        Dim QNo As String
    sqlSelTruck = "SELECT * FROM tbSelTruck" & vbCrLf
    sqlSelTruck = sqlSelTruck & "WHERE QuotationID= '" & tmpRec!QuotationID & "'" & vbCrLf
    Set tmpSel = GetRS(sqlSelTruck, adOpenForwardOnly, adLockReadOnly)
        If tmpSel.RecordCount > 0 Then
            QNo = tmpRec!QuotationNo & "(CR)"
        Else
            QNo = tmpRec!QuotationNo
        End If
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpRec!QuotationID & "_" & tmpRec!wastedataID & "'", QNo)
        tmpList.SubItems(1) = FormatYMD_to_DMY(Trim(tmpRec!IssueDate), "/", "/")
        tmpList.SubItems(2) = tmpRec!CustomerCode
        tmpList.SubItems(3) = Trim(tmpRec!CustomerName)
        tmpList.SubItems(4) = tmpRec!TruckTypeDesc
        tmpList.SubItems(5) = tmpRec!TransportFee
        tmpList.SubItems(6) = tmpRec!TransferUnitID
        If IsNull(tmpRec!CustRatePerTrip) Then
            tmpList.SubItems(7) = "0"
        Else
            tmpList.SubItems(7) = tmpRec!CustRatePerTrip
        End If
        tmpList.SubItems(8) = tmpRec!StaffCountPerCar
        tmpList.SubItems(9) = tmpRec!MinWeightPerCar
        tmpList.SubItems(10) = tmpRec!NonMinRate
        tmpList.SubItems(11) = tmpRec!MaxWeight
        tmpList.SubItems(12) = tmpRec!MaxRate & " " & tmpRec!MaxRateType
        tmpList.SubItems(13) = tmpRec!WasteNo
        tmpList.SubItems(14) = tmpRec!WasteName
        tmpList.SubItems(15) = tmpRec!TreatmentRate
        tmpList.SubItems(16) = tmpRec!TreatmentUnitID
        If tmpRec!IsPriceIncTransport = "Y" Then
            tmpList.SubItems(17) = "รวมค่าขนส่ง"
        Else
            tmpList.SubItems(17) = "แยกค่าขนส่ง"
        End If
        If tmpRec!isCancel = "Y" Then
            tmpList.SubItems(18) = "ยกเลิก"
            tmpList.ForeColor = vbRed
            ForeColor = vbRed
        ElseIf tmpRec!isclosed = "Y" Then
            tmpList.SubItems(18) = "Closed"
            tmpList.ForeColor = &H40C0&
            ForeColor = vbRed
        ElseIf tmpRec!isApproved = "Y" Then
            tmpList.SubItems(18) = "อนุมัติแล้ว"
            tmpList.ForeColor = &H8000&
            ForeColor = &H8000&
        Else
            tmpList.SubItems(18) = "ยังไม่อนุมัติ"
            tmpList.ForeColor = vbBlack
            ForeColor = vbBlack
        End If
                        
        If tmpRec!isJobs = "N" Then
            tmpList.SubItems(19) = "งาน Routine"
        ElseIf tmpRec!isJobs = "Y" Then
            tmpList.SubItems(19) = "งาน Job"
        End If
        
        For i = 1 To tmpList.ListSubItems.Count
    '                    tmpList.ListSubItems(i).Bold = True
            tmpList.ListSubItems(i).ForeColor = ForeColor
        Next
        DoEvents
        Set tmpList = Nothing
        tmpRec.MoveNext
    Loop
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
    Set tmpRec = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
ErrD:
    MsgBox err.Description, vbCritical
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    On Error GoTo ErrD
    Dim tmpRec As New ADODB.Recordset
    Dim backDate As String, tmpArr
    Dim sqlStr As String
    
    Me.Icon = frmMain.Icon
    sqlStr = "select backdate=convert(varchar(10),getdate()-7,111)"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
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
    
    PopulateCompany cboCompany, cboIndex
    
    With cboStatus
        .AddItem "==ทุกสถานะ=="
        .AddItem "ใบเสนอราคาใหม่"
        .AddItem "อนุมัติแล้ว"
        .AddItem "ปิดใบเสนอราคา"
        .AddItem "ยกเลิก"
        .ListIndex = 0
    End With
    
    Exit Sub
    
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error Resume Next
    With lvResult
        If Me.Width - Picture1.Width > 500 Then
            .Width = Me.Width - Picture1.Width - 200
        End If
        .Height = Me.Height - .Top - 500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VisibleImage
End Sub

Private Sub ClearHeaderIcons(CurrentHeader As Integer)
    Dim i As Integer
    For i = 1 To lvResult.ColumnHeaders.Count
        If lvResult.ColumnHeaders(i).Index <> CurrentHeader Then
            lvResult.ColumnHeaders(i).Icon = Empty
        End If
    Next
End Sub

Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call ClearHeaderIcons(ColumnHeader.Index)
    Select Case ColumnHeader.Index
        Case 2
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortDescending, sortDate)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortDate)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortDate)
            End Select
        Case Else
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortDescending, sortAlpha)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortAlpha)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortAlpha)
            End Select
    End Select
End Sub

Private Sub lvResult_DblClick()
    Dim SplitID() As String
    Dim tmpS  As BWGSearchManager
    SplitID = Split(lvResult.SelectedItem.Key, "_")
    If SplitID(0) = "" Or SplitID(0) = "''" Then Exit Sub
    Set tmpS = New BWGSearchManager
    Set colSearchResult_ = tmpS.QuotationHeaderSearch("QuotationID = " & SplitID(0) & "'")
    frmQuotation.SetDataQuotation colSearchResult_(1)
    frmQuotation.Show 1
End Sub

Private Sub TxtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCustomer_Validate False
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "*" Or Trim(txtCustomer.Text) = "" Then Set curCustomer_ = Nothing: Exit Sub
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .GetCriteriaOptonal " AND CustGroupID <> '488F2E5C057F46869D1A56185B11D644'"
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_.ID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCust As BWGCustomer, Optional isAssigned As Boolean = True)
    If selCust Is Nothing Then
        txtCustomer.Text = ""
        Exit Sub
    End If
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCust
    With curCustomer_
        txtCustomer.Text = .CustomerName
    End With
End Sub

Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
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

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
    Set SaleStaff_ = Nothing
    Set SaleStaff_ = selStaff
    If SaleStaff_ Is Nothing Then Exit Sub
    With SaleStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub
