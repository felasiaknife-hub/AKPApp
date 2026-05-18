VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmPrintTRN13 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   6105
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CbSaleType 
      Height          =   315
      ItemData        =   "frPrintTRN13.frx":0000
      Left            =   1350
      List            =   "frPrintTRN13.frx":0002
      TabIndex        =   30
      Top             =   5040
      Width           =   1950
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4635
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Width           =   5985
      Begin VB.ComboBox cboCustType 
         Height          =   315
         ItemData        =   "frPrintTRN13.frx":0004
         Left            =   1140
         List            =   "frPrintTRN13.frx":0006
         TabIndex        =   28
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   25
         Top             =   3480
         Width           =   4395
      End
      Begin VB.TextBox txtVendor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   23
         Top             =   1140
         Width           =   4425
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3120
         Width           =   4425
      End
      Begin VB.TextBox txtFromAumphur 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Top             =   2760
         Width           =   1875
      End
      Begin VB.TextBox txtFromProvince 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2760
         Width           =   1755
      End
      Begin VB.ComboBox cboShowRecords 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1980
         Width           =   4425
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   2400
         Width           =   4395
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   750
         Width           =   4425
      End
      Begin VB.ComboBox cboCarType 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1560
         Width           =   4425
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   1
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทการขาย"
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   29
         Top             =   4230
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทลูกค้า"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   27
         Top             =   3885
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         Height          =   195
         Index           =   8
         Left            =   600
         TabIndex        =   26
         Top             =   3540
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor"
         Height          =   195
         Index           =   7
         Left            =   480
         TabIndex        =   24
         Top             =   1260
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ออเดอร์"
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   22
         Top             =   3180
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ต้นทาง อำเภอ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   2820
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "จังหวัด"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   19
         Top             =   2820
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "************"
         Height          =   195
         Index           =   5
         Left            =   300
         TabIndex        =   18
         Top             =   2100
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   17
         Top             =   2460
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ขนโดย"
         Height          =   195
         Index           =   4
         Left            =   540
         TabIndex        =   16
         Top             =   810
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   15
         Top             =   1620
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่ขน"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   14
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3120
         TabIndex        =   13
         Top             =   420
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5505
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5505
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5535
      Width           =   915
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   120
      X2              =   5760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   5910
      Y1              =   570
      Y2              =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานการขนกากแยกตามประเภทรถ (TRN13)"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   60
      Width           =   4050
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frPrintTRN13.frx":0008
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintTRN13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff
Public isMD As Boolean
Dim curFromAumphur_ As IMIAumphur
Dim curToAumphur_ As IMIAumphur
Dim curAumphurType_ As String
Dim curVendor_ As BWGTransportVendor
Dim curCustomer_ As BWGCustomer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    If cboCarType.ListIndex = -1 Then
        MsgBox "กรุณาระบุประเภทรถ", vbExclamation
        cboCarType.SetFocus
        Exit Sub
    End If
    PrintMain 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_.ID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub
Private Sub cmdPrint_Click()
    If cboCarType.ListIndex = -1 Then
        MsgBox "กรุณาระบุประเภทรถ", vbExclamation
        cboCarType.SetFocus
        Exit Sub
    End If
    PrintMain 1
End Sub

Private Sub Form_Load()
    PopulateCarType cboCarType
    PopulateCompany cboCompany, cboIndex
    With cboShowRecords
        .AddItem "==แสดงทุกรายการ=="
        .AddItem "แสดงเฉพาะรายการที่ค่าขนส่งแยก น้อยกว่า ค่าขนส่งที่จ่ายเวนเดอร์"
        .AddItem "แสดงเฉพาะรายการที่ค่าขนส่งแฝง = 0"
        .ListIndex = 0
    End With
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    
    With cboOrderType
        .AddItem "ออร์เดอร์แบบปกติ (Routine)"
        .AddItem "ออร์เดอร์แบบ Jobs"
        .AddItem "ออร์เดอร์น้ำท่วม Mixed Waste"
        .AddItem "ทั้งหมด"
        .ListIndex = 0
    End With
        With cboCustType
        .AddItem "== ประเภทลูกค้า =="
        .AddItem "(C) ,(s)"
        .AddItem "(M) ,(MS)"
        .AddItem "(B) ,(E)"
        .ListIndex = 0
    End With
    With CbSaleType
        .AddItem "ขายตรง"
        .AddItem "ขายผ่าน BME"
        .AddItem "ทั้งหมด"
        .ListIndex = 0
    End With
    dtFrom.DayValue = 1
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If isMD = False Then
            .ReportFileName = App.Path & "\reports\TRN13.rpt"
        Else
            .ReportFileName = App.Path & "\reports\TRN13_MD.rpt"
        End If
        tmpSql = "SELECT * FROM vw_MenifestSearch" & vbCrLf
        tmpSql = tmpSql & "WHERE WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "'" & vbCrLf
        tmpSql = tmpSql & "AND isCanceled <> 'Y' AND isClosed = 'Y'" & vbCrLf
'        tmpSql = tmpSql & "AND ActualQty > 0" & vbCrLf

'        tmpSql = tmpSql & "AND CustomerName LIKE '%นวนคร%')" & vbCrLf
'        tmpSql = tmpSql & "OR (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "'" & vbCrLf
'        tmpSql = tmpSql & "AND isCanceled <> 'Y' AND isClosed = 'Y'" & vbCrLf
'        tmpSql = tmpSql & "AND SiteAumphurName = 'คลองหลวง')" & vbCrLf
        
        If cboCarType.ListIndex > 0 Then tmpSql = tmpSql & "AND TruckTypeID='" & coltruckType_(cboCarType.ListIndex).ID & "'" & vbCrLf
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
            End If
        End If
          If Not curCustomer_ Is Nothing Then
            If Trim(curCustomer_.ID) <> "" Then
                tmpSql = tmpSql & " And CustomerID='" & curCustomer_.ID & "' "
            End If
        End If
        If Trim(txtSale.Text) <> "" Then
            tmpSql = tmpSql & " And UnderSaleStaffID='" & curStaff_.ID & "' " & vbCrLf
        End If
        
        If Not curFromAumphur_ Is Nothing Then
            tmpSql = tmpSql & "AND SiteAumphurID = '" & curFromAumphur_.ID & "'" & vbCrLf
        End If
        
        If Trim(txtVendor.Text) <> "" Then
               tmpSql = tmpSql & "AND VendorID = '" & curVendor_.ID & "'" & vbCrLf
        End If
        If cboCustType.ListIndex = 1 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('C','S') "
        ElseIf cboCustType.ListIndex = 2 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('M','MS') "
        ElseIf cboCustType.ListIndex = 3 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('B','E') "
        End If

        If cboOrderType.ListIndex = 0 Then
'            tmpSql = tmpSql & "AND CustomerID IN (SELECT CustomerID FROM tbCustomer WHERE CustomerCode IN (SELECT CustomerCode FROM tbCustomer_MixedWaste))" & vbCrLf
             tmpSql = tmpSql & " And  isjobs <> 'Y' And (WasteCaseID is null or WasteCaseID = '')" & vbCrLf
        ElseIf cboOrderType.ListIndex = 1 Then
             tmpSql = tmpSql & " And isJobs = 'Y' " & vbCrLf
        ElseIf cboOrderType.ListIndex = 2 Then
             tmpSql = tmpSql & " And  WasteCaseID is not null  And WasteCaseID <> ''" & vbCrLf
        End If
        
        If CbSaleType.ListIndex = 0 Then
            tmpSql = tmpSql & " And  refme is null " & vbCrLf
        ElseIf CbSaleType.ListIndex = 1 Then
            tmpSql = tmpSql & " And  refme is not null " & vbCrLf
        End If
        
'        tmpSql = tmpSql & "AND CustomerID IN (SELECT CustomerID FROM tbCustomer WHERE DTeam = 'Y')"
'        tmpSql = tmpSql & "AND CustomerID IN (SELECT CustomerID FROM vw_CustomerAddress WHERE AumphurName LIKE '%นครหลวง%' AND TumbolName = 'บางพระครู')" & vbCrLf
        
        tmpSql = tmpSql & "ORDER BY CompanyNameTH, TruckTypeDesc, CustomerName, TimeTableNo, IncollectFee Desc"
        
        .Formulas(0) = "DateRange=" & Chr(34) & "ในช่วงวันที่ขนตั้งแต่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .Formulas(2) = "ShowCompany=" & Chr(34) & cboCompany.Text & Chr(34)
        .Formulas(3) = "ReportTitle = " & Chr(34) & cboCarType.Text & Chr(34)
        If cboCarType.ListIndex <= 0 Then .Formulas(4) = "ShowGroup1 = " & Chr(34) & "False" & Chr(34)
        If cboShowRecords.ListIndex > 0 Then
            .Formulas(5) = "ShowRecords = " & Chr(34) & cboShowRecords.ListIndex & Chr(34)
            .ReportFileName = App.Path & "\reports\TRN13_MD.rpt"
        End If
        If cboOrderType.ListIndex = 2 Then
            .ReportFileName = App.Path & "\reports\TRN13_Waste.rpt"
        End If
        .WindowTitle = "รายงานการขนกากแยกตามประเภทรถ (TRN13) - (เวสต์น้ำท่วม - Mixed Waste)"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub txtFromAumphur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtFromAumphur_Validate False
End Sub

Private Sub txtFromAumphur_Validate(Cancel As Boolean)
    If curFromAumphur_ Is Nothing Then Set curFromAumphur_ = New IMIAumphur
    If Trim(txtFromAumphur.Text) = "" And Trim(curFromAumphur_.Name) <> Trim(txtFromAumphur.Text) Then
        Set curFromAumphur_ = Nothing
        Set curFromAumphur_ = New IMIAumphur
        txtFromProvince.Text = ""
    Else
        If Trim(curFromAumphur_.Name) <> Trim(txtFromAumphur.Text) Then
            Set curFromAumphur_ = Nothing
            Set curFromAumphur_ = New IMIAumphur
            txtFromProvince.Text = ""
            Load frmAumphurSearch
            With frmAumphurSearch
                .getCriteria (txtFromAumphur.Text)
                .SetCallerForm Me
                .Show 1
            End With
        End If
    End If
    curAumphurType_ = "F"
    If txtFromAumphur.Text = "" Then Set curFromAumphur_ = Nothing
End Sub

Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtSale.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curStaff_.ID = "" Then txtSale.Text = ""
        End If
    End If
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curStaff_ = Nothing
    Set curStaff_ = selStaff
    If curStaff_ Is Nothing Then Exit Sub
    With curStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub

Public Sub SetData(selAumphur As IMIAumphur, Optional isAssigned As Boolean = True)
    If curAumphurType_ = "F" Then
        Set curFromAumphur_ = Nothing
        Set curFromAumphur_ = selAumphur
        If curFromAumphur_ Is Nothing Then Exit Sub
        With curFromAumphur_
            txtFromAumphur.Text = .Name
            txtFromProvince.Text = .ProvinceName
        End With
    ElseIf curAumphurType_ = "T" Then
        Set curToAumphur_ = Nothing
        Set curToAumphur_ = selAumphur
        If curToAumphur_ Is Nothing Then Exit Sub
        With curToAumphur_
'            txtToAumphur.Text = .Name
'            txtToProvince.Text = .ProvinceName
        End With
    End If
End Sub

Private Sub txtVendor_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then txtVendor_Validate False
End Sub

Private Sub txtVendor_Validate(Cancel As Boolean)
Dim selStr As String
    If curVendor_ Is Nothing Then Set curVendor_ = New BWGTransportVendor
    selStr = Trim(curVendor_.NameTH)
    If Trim(txtVendor.Text) = "" And Trim(selStr) <> Trim(txtVendor.Text) Then
        Set curVendor_ = Nothing
        Set curVendor_ = New BWGTransportVendor
    Else
        If Trim(selStr) <> Trim(txtVendor.Text) Then
            Set curVendor_ = Nothing
            Set curVendor_ = New BWGTransportVendor
            Load frmTransportVenderSearch
            With frmTransportVenderSearch
                .getCriteria (txtVendor.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curVendor_ Is Nothing Then
                txtVendor.Text = ""
            Else
                If curVendor_.ID = "" Then
                    txtVendor.Text = ""
                End If
            End If
        End If
    End If
End Sub
Public Sub SetDataTransportVender(selVendor As BWGTransportVendor, Optional isAssigned As Boolean = True)
    Set curVendor_ = Nothing
    Set curVendor_ = selVendor
    If curVendor_ Is Nothing Then Exit Sub
    With curVendor_
        txtVendor.Text = .NameTH
    End With
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
Dim tmpWaste As BWGWasteDataCR
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    With curCustomer_
        txtCustomer.Text = .CustomerName
    End With
End Sub
