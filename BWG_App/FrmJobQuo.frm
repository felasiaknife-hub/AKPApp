VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmJobQuo 
   Caption         =   "ตรวจสอบ Job"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   14220
   ForeColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   14220
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtPriceTo 
      Height          =   285
      Left            =   1575
      TabIndex        =   21
      Top             =   1665
      Width           =   915
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2640
      TabIndex        =   19
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImgSorted"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Job No."
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่ขน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ลูกค้า"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "จังหวัด"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "พนักงานขาย"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "เลขที่ใบเสนอราคา"
         Object.Width           =   2795
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ชื่อ Waste"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ค่าบริการ"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "หน่วย"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ประเภทค่าขนส่ง"
         Object.Width           =   2690
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "รถขนส่ง"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ค่าขนส่ง"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "JobID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "น้ำหนัก"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "IsNewWaste"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8715
      Left            =   0
      ScaleHeight     =   8685
      ScaleWidth      =   2625
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   990
         TabIndex        =   28
         Top             =   4635
         Width           =   1500
      End
      Begin VB.CheckBox Check1 
         Caption         =   "ดูข้อมูลทั้งหมด (ถ้าต้องการดูเฉพาะ บาท/ตัน ไม่ต้องเลือก)"
         Height          =   555
         Left            =   90
         TabIndex        =   27
         Top             =   7065
         Width           =   2400
      End
      Begin VB.ComboBox cboCarType 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   2295
         Width           =   2310
      End
      Begin VB.CommandButton CmdExp 
         Caption         =   "Excel"
         Height          =   435
         Left            =   180
         TabIndex        =   24
         Top             =   6075
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox TxtPrice 
         Height          =   285
         Left            =   180
         TabIndex        =   20
         Top             =   1665
         Width           =   960
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "พิมพ์รายงาน"
         Height          =   435
         Left            =   180
         TabIndex        =   5
         Top             =   5580
         Width           =   2295
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   180
         TabIndex        =   4
         Top             =   6570
         Width           =   2295
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   180
         TabIndex        =   3
         Top             =   5085
         Width           =   2295
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "FrmJobQuo.frx":0000
         Left            =   180
         List            =   "FrmJobQuo.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2925
         Width           =   2325
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Text            =   "*"
         Top             =   360
         Width           =   2295
      End
      Begin DateCtl.Date dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   7
         Top             =   3600
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin DateCtl.Date dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   8
         Top             =   4215
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อ Waste :"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   29
         Top             =   4635
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "ประเภทรถ :"
         Height          =   195
         Left            =   225
         TabIndex        =   26
         Top             =   2070
         Width           =   840
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "ราคาระหว่าง"
         Height          =   195
         Left            =   225
         TabIndex        =   23
         Top             =   1395
         Width           =   1005
      End
      Begin VB.Line Line2 
         X1              =   180
         X2              =   2490
         Y1              =   4995
         Y2              =   4995
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย :"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   18
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   450
         TabIndex        =   17
         Top             =   9120
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปิดใบเสนอราคา"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   450
         TabIndex        =   16
         Top             =   8820
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อนุมัติแล้ว"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   450
         TabIndex        =   15
         Top             =   8520
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "งานใหม่"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   450
         TabIndex        =   14
         Top             =   8220
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   180
         Top             =   9090
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   180
         Top             =   8790
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   180
         Top             =   8490
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   180
         Top             =   8190
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   13
         Top             =   3975
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ขน :"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   12
         Top             =   3360
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสหรือชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   11
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ค่าขนส่ง :"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   2685
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึง"
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   9
         Top             =   1710
         Width           =   180
      End
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   4980
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
            Picture         =   "FrmJobQuo.frx":0004
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmJobQuo.frx":059E
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ราคาต่ำกว่า"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   22
      Top             =   1710
      Width           =   795
   End
End
Attribute VB_Name = "FrmJobQuo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaleStaff_ As BWGCompanyStaff
Dim curCustomer_ As BWGCustomer
Dim colSearchResult_ As Collection

Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Function SQLString() As String
    Dim strCond As String
    Dim criStr As String
    criStr = Replace(Replace(Trim(txtCustomer.Text), "*", "%"), "'", "")
'    strCond = "(QuotationNo Like '%" & criStr & "%' Or CustomerCode Like '%" & criStr & "%' Or CompanyName Like '%" & criStr & "%')"
    strCond = "(workDate >= '" & dtFrom.ValueYMD & "' AND workdate<= '" & dtTo.ValueYMD & "')"
    If Not curCustomer_ Is Nothing Then strCond = strCond & vbCrLf & "AND CustomerID = '" & curCustomer_.ID & "'"
    If Not SaleStaff_ Is Nothing Then strCond = strCond & vbCrLf & "AND (UnderSaleStaffID = '" & SaleStaff_.ID & "')"
    
    If cboStatus.ListIndex = 1 Then
        strCond = strCond & vbCrLf & "AND ISNULL(isPriceIncTransport, 'N') = 'Y'"
    ElseIf cboStatus.ListIndex = 2 Then
        strCond = strCond & vbCrLf & "AND ISNULL(isPriceIncTransport, 'N') = 'N'"
    End If
    If Trim(TxtPrice.Text) <> "" And Trim(TxtPriceTo.Text) <> "" Then
        strCond = strCond & vbCrLf & "AND treatmentrate>= '" & Trim(TxtPrice.Text) & "' and treatmentrate<= '" & Trim(TxtPriceTo.Text) & "'"
    ElseIf Trim(TxtPrice.Text) <> "" And TxtPriceTo.Text = "" Then
        strCond = strCond & vbCrLf & "AND treatmentrate>= '" & Trim(TxtPrice.Text) & "' "
    End If
    If cboCarType.ListIndex > 0 Then strCond = strCond & "AND TruckTypeID='" & coltruckType_(cboCarType.ListIndex).ID & "'" & vbCrLf
    If Check1.Value = 0 Then strCond = strCond & "AND TreatmentUnitID='บาท/ตัน' " & vbCrLf

        strCond = strCond & vbCrLf & " AND iscancel<>'Y' and JobNo<>'' and customercode<>'C02300' "
    
    SQLString = strCond
End Function

Private Sub CmdExp_Click()
'    Dim objExcel As New Excel.Application
'
'    Dim objExcelSheet As Excel.Worksheet
'    Dim col, row As Integer
'
'    If lvResult.ListItems.Count > 0 Then
'        objExcel.Workbooks.Add
'        Set objExcelSheet = objExcel.Worksheets.Add
'
'
'        For col = 1 To lvResult.ColumnHeaders.Count
'            objExcelSheet.Cells(1, col).Value = lvResult.ColumnHeaders(col)
'        Next
'
'        For row = 2 To lvResult.ListItems.Count + 1
'           For col = 1 To lvResult.ColumnHeaders.Count
'           If col = 1 Then
'                   objExcelSheet.Cells(row, col).Value = lvResult.ListItems(row - 1).Text
'           Else
'                   objExcelSheet.Cells(row, col).Value = lvResult.ListItems(row - 1).SubItems(col - 1)
'           End If
'           Next
'        Next row
'
'        objExcelSheet.Columns.AutoFit
'
'
'        objExcel.Visible = True
'
'    Else
'        MsgBox "No data to export", vbInformation, Me.Caption
'    End If

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
        .WindowTitle = "Job Quotation Detail"
        .ReportFileName = App.Path & "\Reports\rptcheckorder.rpt"
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If TxtPrice.Text <> "" And TxtPriceTo.Text <> "" Then
            .Formulas(0) = "dtfrom = " & Chr(34) & "ราคาระหว่าง" & TxtPrice & " ถึง " & TxtPriceTo & Chr(34)
        ElseIf TxtPrice.Text <> "" And TxtPriceTo.Text = "" Then
            .Formulas(0) = "dtfrom = " & Chr(34) & "ราคาตั้งแต่" & TxtPrice & " ขึ้นไป " & Chr(34)
        End If
        tmpSql = "SELECT * from TbRptCheckJob"
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
    If Trim(Text1.Text) <> "" Then
        strCond = strCond & " and wastename like '%" & Trim(Text1.Text) & "%' "
    End If

'    strCond = SelectWithCompanyID(strCond, True)
'    Set colSearchResult_ = tmpS.QuotationDetailSearch(strCond)
'    Set tmpS = Nothing
'    Set tmpItem = colSearchResult_(x)
'    Set tmpColDet = tmpRec!QuotationDetails
    sqlStr = "SELECT JobDetID,JobID,JobNo,WorkDate,CustomerCode,CompanyName,Addr1,TumbolName,AumphurName,ProvinceName,SaleName,QuotationNo,WasteName,TreatmentRate,TreatmentUnitID," & vbCrLf
    sqlStr = sqlStr & "MinWeightPerCar,isPriceIncTransport,isCancel,TruckTypeDesc,CustomerID,UnderSaleStaffID,sum(disposerweight) as Disposerweight,transportfee,isNewWaste  FROM vw_JobQuotationDetail" & vbCrLf
    sqlStr = sqlStr & "WHERE " & strCond & vbCrLf
    sqlStr = sqlStr & "group by JobDetID,JobID,JobNo,WorkDate,CustomerCode,CompanyName,Addr1,TumbolName,AumphurName,ProvinceName,SaleName,QuotationNo,WasteName,TreatmentRate,TreatmentUnitID," & vbCrLf
    sqlStr = sqlStr & "MinWeightPerCar,isPriceIncTransport,isCancel,TruckTypeDesc,CustomerID,UnderSaleStaffID,transportfee,isNewWaste ORDER BY  workdate, QuotationNo" & vbCrLf
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    
    If tmpRec.RecordCount = 0 Then Screen.MousePointer = 0: Exit Sub
        DBConnExc "Delete TbRptCheckJob "

    frmMain.Stb1.Panels(1).Text = "พบข้อมูลทั้งหมด " & Format(tmpRec.RecordCount, "#,##0") & " รายการ"
    Screen.MousePointer = 0
    DoEvents
    Screen.MousePointer = 11
    Do While Not tmpRec.EOF
        Dim ForeColor
        Dim Job As String
        If tmpRec!JobNo <> "" Then
            Job = tmpRec!JobNo
            Else
            Job = "-"
        End If
        
   DBConnExc "Insert into TbRptCheckJob (JobDetID,JobID,JobNo,WorkDate,CustomerCode,CompanyName,Addr1,TumbolName,AumphurName " & _
                            " ,ProvinceName,SaleName,QuotationNo,WasteName,TreatmentRate,TreatmentUnitID,MinWeightPerCar,isPriceIncTransport,isCancel " & _
                            ",TruckTypeDesc,CustomerID,UnderSaleStaffID,Disposerweight,transportfee,isNewWaste) " & _
                            " Values ('" & tmpRec!JobDetID & "','" & tmpRec!JobID & "','" & tmpRec!JobNo & "','" & tmpRec!WorkDate & "','" & tmpRec!CustomerCode & "' " & _
                            ",'" & tmpRec!CompanyName & "','" & tmpRec!Addr1 & "','" & tmpRec!TumbolName & "','" & tmpRec!AumphurName & "','" & tmpRec!ProvinceName & "' " & _
                            ",'" & tmpRec!salename & "','" & tmpRec!QuotationNo & "','" & tmpRec!WasteName & "','" & tmpRec!TreatmentRate & "','" & tmpRec!TreatmentUnitID & "' " & _
                            ",'" & tmpRec!MinWeightPerCar & "','" & tmpRec!IsPriceIncTransport & "','" & tmpRec!isCancel & "','" & tmpRec!TruckTypeDesc & "','" & tmpRec!CustomerID & "' " & _
                            ",'" & tmpRec!UnderSaleStaffID & "','" & tmpRec!DisposerWeight & "','" & tmpRec!TransportFee & "','" & tmpRec!IsNewWaste & "')"
        
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpRec!JobDetID & "_" & tmpRec!JobID & "'", Job)
        tmpList.SubItems(1) = FormatYMD_to_DMY(Trim(tmpRec!WorkDate), "/", "/")
        tmpList.SubItems(2) = "(" & Trim(tmpRec!CustomerCode) & ")  " & Trim(tmpRec!CompanyName)
        tmpList.SubItems(3) = tmpRec!ProvinceName
        If tmpRec!salename <> "Null" Then
        tmpList.SubItems(4) = tmpRec!salename
        End If
        tmpList.SubItems(5) = tmpRec!QuotationNo
        tmpList.SubItems(6) = tmpRec!WasteName
        tmpList.SubItems(7) = tmpRec!TreatmentRate
        tmpList.SubItems(8) = tmpRec!TreatmentUnitID
        If tmpRec!IsPriceIncTransport = "Y" Then
            tmpList.SubItems(9) = "รวมค่าขนส่ง"
        Else
            tmpList.SubItems(9) = "แยกค่าขนส่ง"
        End If
        tmpList.SubItems(10) = tmpRec!TruckTypeDesc
        tmpList.SubItems(11) = tmpRec!TransportFee
        tmpList.SubItems(12) = tmpRec!JobID
        tmpList.SubItems(13) = tmpRec!DisposerWeight
        If IsNull(tmpRec!IsNewWaste) = True Then
        tmpList.SubItems(14) = ""
        Else
        tmpList.SubItems(14) = tmpRec!IsNewWaste
        End If
'        If tmpRec!isCancel = "Y" Then
'            tmpList.SubItems(18) = "ยกเลิก"
'            tmpList.ForeColor = vbRed
'            ForeColor = vbRed
'        ElseIf tmpRec!isclosed = "Y" Then
'            tmpList.SubItems(18) = "Closed"
'            tmpList.ForeColor = &H40C0&
'            ForeColor = vbRed
'        ElseIf tmpRec!isApproved = "Y" Then
'            tmpList.SubItems(18) = "อนุมัติแล้ว"
'            tmpList.ForeColor = &H8000&
'            ForeColor = &H8000&
'        Else
'            tmpList.SubItems(18) = "ยังไม่อนุมัติ"
'            tmpList.ForeColor = vbBlack
'            ForeColor = vbBlack
'        End If
        
'        For i = 1 To tmpList.ListSubItems.Count
'    '                    tmpList.ListSubItems(i).Bold = True
'            tmpList.ListSubItems(i).ForeColor = ForeColor
'        Next
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

Private Sub Form_Resize()
On Error Resume Next
    With lvResult
        If Me.Width - Picture1.Width > 500 Then
            .Width = Me.Width - Picture1.Width - 200
        End If
        .Height = Me.Height - .Top - 500
    End With
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
        
    With cboStatus
        .AddItem "==ทั้งหมด=="
        .AddItem "รวมค่าขนส่ง"
        .AddItem "แยกค่าขนส่ง"
        .ListIndex = 0
    End With
        PopulateCarType cboCarType

    Exit Sub
    
ErrD:
    GetMsg "", vbCritical
    Exit Sub

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

Private Sub lvResult_DblClick()
On Error GoTo ErrD
Dim tmpS As New BWGSearchManager, tmpCol As Collection
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    Load frmJobOpen
    With frmJobOpen
        Set tmpCol = tmpS.JobDataHeaderSearch("JobID= '" & lvResult.SelectedItem.SubItems(12) & "'")
        If tmpCol.Count > 0 Then
            .SetDataJobOpen tmpCol(1)
            Set tmpCol = Nothing
            Set tmpS = Nothing
            .Show 1
        End If
    End With
    Set tmpCol = Nothing
    Set tmpS = Nothing
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub

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

