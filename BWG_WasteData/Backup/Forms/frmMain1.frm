VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain1 
   Caption         =   "Preliminary Industrial - Waste Data Form"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   13170
   Icon            =   "frmMain1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   13170
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   8580
      TabIndex        =   19
      Top             =   2940
      Width           =   1395
   End
   Begin VB.Frame Frame2 
      Caption         =   "2. ส่วนของข้อมูลโรงงาน : Company Data"
      Height          =   3195
      Left            =   60
      TabIndex        =   6
      Top             =   2220
      Width           =   13095
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6600
         TabIndex        =   38
         Top             =   2760
         Width           =   2475
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8160
         TabIndex        =   35
         Top             =   2280
         Width           =   4155
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1200
         TabIndex        =   33
         Top             =   2280
         Width           =   5295
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8160
         TabIndex        =   31
         Top             =   1860
         Width           =   4155
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1200
         TabIndex        =   30
         Top             =   1860
         Width           =   5295
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   8160
         TabIndex        =   27
         Top             =   1440
         Width           =   4155
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1860
         TabIndex        =   26
         Top             =   1440
         Width           =   4635
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   11100
         TabIndex        =   23
         Top             =   1020
         Width           =   1155
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6480
         TabIndex        =   18
         Top             =   1020
         Width           =   1395
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4140
         TabIndex        =   17
         Top             =   1020
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1200
         TabIndex        =   9
         Top             =   1020
         Width           =   2355
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1620
         TabIndex        =   7
         Text            =   "*"
         Top             =   300
         Width           =   7755
      End
      Begin VB.Label Label18 
         Caption         =   "เลข 13 หลัก (ถ้ามี)"
         Height          =   195
         Left            =   9120
         TabIndex        =   39
         Top             =   2820
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "5. เลขประจำตัวผู้ดำเนินการเกี่ยวกับการจัดการของเสียอันตราย (ผู้ก่อกำเนิดของเสียอันตราย)"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   2820
         Width           =   6495
      End
      Begin VB.Label Label16 
         Caption         =   "แฟกซ์"
         Height          =   195
         Left            =   7620
         TabIndex        =   36
         Top             =   2340
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "โทรศัพท์"
         Height          =   195
         Left            =   480
         TabIndex        =   34
         Top             =   2340
         Width           =   675
      End
      Begin VB.Label Label14 
         Caption         =   "แผนก"
         Height          =   195
         Left            =   7620
         TabIndex        =   32
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label13 
         Caption         =   "4. ชื่อผู้ติดต่อ :"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "เลขที่ทะเบียนโรงงาน"
         Height          =   195
         Left            =   6600
         TabIndex        =   28
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label11 
         Caption         =   "3. ประเภทอุตสาหกรรม :"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   1500
         Width           =   1755
      End
      Begin VB.Label Label10 
         Caption         =   "รหัสไปรษณีย์"
         Height          =   195
         Left            =   10080
         TabIndex        =   24
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "จังหวัด"
         Height          =   195
         Left            =   7980
         TabIndex        =   22
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "อำเภอ"
         Height          =   195
         Left            =   6000
         TabIndex        =   21
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label7 
         Caption         =   "ตำบล"
         Height          =   195
         Left            =   3660
         TabIndex        =   20
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label6 
         Caption         =   "เลขที่, หมู่"
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2. ที่อยู่บริษัท/โรงงาน :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1. ชื่อบริษัท/โรงงาน :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "3. ส่วนข้อมูลของเสียอุตสาหกรรม : Waste Data"
      Height          =   1935
      Left            =   60
      TabIndex        =   1
      Top             =   6000
      Width           =   13095
   End
   Begin VB.Frame Frame1 
      Caption         =   "1. ส่วนของ บริษัท เบตเตอร์ เวิลด์ กรีน จำกัด (มหาชน) : This section is for Better World Green Public Company Limited."
      Height          =   1395
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   13095
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         TabIndex        =   12
         Text            =   "*"
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1740
         TabIndex        =   11
         Text            =   "*"
         Top             =   360
         Width           =   4575
      End
      Begin VB.TextBox txtWaste 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   4440
         Width           =   2235
      End
      Begin Crystal.CrystalReport rpt1 
         Left            =   120
         Top             =   7380
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "(เลข 6 หลัก)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6480
         TabIndex        =   15
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "รหัสของเสีย"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "เลขที่ของเสีย"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อของเสีย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   4
         Top             =   4200
         Width           =   720
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7890
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
         EndProperty
      EndProperty
   End
   Begin WasteDataForm.ctlDate dtFrom 
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   60
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   661
      BackColor       =   -2147483633
   End
   Begin VB.Label Label1 
      Caption         =   "วันที่"
      Height          =   195
      Left            =   7320
      TabIndex        =   40
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMain1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff
Dim colZone_ As Collection
Dim sessionID_ As String
Dim curCustomer_ As BWGCustomer
Dim Counts As Long
Dim SaleStaffID As String
Dim CRStaffID As String

Private Sub cmdPrint_Click()
    PrintMain 1
End Sub

Private Sub cmdProgress_Click()
On Error GoTo ErrD
Dim RS As ADODB.Recordset
Dim tSql As String
Dim i As Integer
    If Result.ListItems.Count <= 0 Then
        MsgBox "ไม่มีข้อมูล", vbInformation, "Warning"
        Exit Sub
    End If
    If MsgBox("ยืนยันการเปลี่ยนแปลงข้อมูล", vbExclamation + vbYesNo, "Comfirm") = vbNo Then Exit Sub
    DBManager.GetDBConnection.BeginTrans
    For i = 1 To Result.ListItems.Count
        If Result.ListItems(i).Checked = True Then
            tSql = "UPDATE tbCustomer SET UnderSaleStaffID = '', SaleDistrictID = ''" & vbCrLf
            tSql = tSql & "WHERE CustomerID = '" & Result.ListItems(i).SubItems(3) & "'"
            DBManager.GetDBConnection.Execute tSql
            
            tSql = "UPDATE Com_vwQuotation SET isTreatAble = 'N', isApproveByLab = 'R'" & vbCrLf
            tSql = tSql & "WHERE CustomerID = '" & Result.ListItems(i).SubItems(3) & "'" & vbCrLf
            tSql = tSql & "AND IssueDate >= '" & frmMain.dtFrom.ValueYMD & "'" & vbCrLf
            tSql = tSql & "AND IssueDate <= '" & frmMain.dtTo.ValueYMD & "'"
            DBManager.GetDBConnection.Execute tSql
        End If
    Next
    DBManager.GetDBConnection.CommitTrans
    MsgBox "ประมวลผลเสร็จแล้ว", vbInformation, "Completed"
    Exit Sub
ErrD:
    DBManager.GetDBConnection.RollbackTrans
    Exit Sub
End Sub

Private Sub cmdSearch_Click()
Dim RS As ADODB.Recordset
Dim RS2 As ADODB.Recordset
Dim tSql As String
Dim Query1 As String
Dim Query2 As String
Dim Cn As Long
    Result.ListItems.Clear
        Query1 = "ใบเสนอราคา"
        tSql = SqlQuery("Com_vwQuotationSearch", 0)
        Set RS = New ADODB.Recordset
        
        StatusBar1.Panels(1).Text = "กำลังค้นหา " & Query1
        DoEvents
        RS.Open tSql, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
        DoEvents
        StatusBar1.Panels(1).Text = "พบข้อมูล " & Query1 & " จำนวน " & RS.RecordCount & " รายการ"
'        Load frmProgress
'        frmProgress.Show 1
        
        Call AddtoListView(RS, 1)
        Set RS = Nothing
End Sub

Private Sub Form_Activate()
    Call ActivateForm
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Call ActivateForm
End Sub

Private Sub ActivateForm()
On Error Resume Next
'    Frame1.Height = Me.Height - StatusBar1.Height - 350
    
'    Frame2.Height = Me.Height - StatusBar1.Height - 350
    Label1.Left = Me.Width - (dtFrom.Width + 200 + Label1.Width)
    dtFrom.Left = Me.Width - (dtFrom.Width + 200)
    Frame1.Width = Me.Width - Frame1.Left - 150
    Frame2.Width = Me.Width - Frame2.Left - 150
    Frame3.Width = Me.Width - Frame3.Left - 150


'    dtFrom.ValueYMD = Left(TodayDate, 8) & "01"
'    dtTo.ValueYMD = TodayDate
    
    StatusBar1.Height = 435
    StatusBar1.Panels(1).Width = Me.Width * 30 / 100
    StatusBar1.Panels(2).Width = Me.Width * 70 / 100
End Sub

Private Sub cmdClose_Click()
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub

Private Sub cmdPreview_Click()
Dim tSql As String
Dim i As Integer
'    tSql = "DELETE FROM Com_tbTmpCustomer"
'    DBManager.GetDBConnection.Execute tSql
'    For i = 1 To Result.ListItems.Count
'        If Result.ListItems(i).Checked = True Then
'            tSql = "INSERT INTO Com_tbTmpCustomer(CustomerID)"
'            tSql = tSql & "VALUES ('" & Result.ListItems(i).SubItems(3) & "')"
'            DBManager.GetDBConnection.Execute tSql
'        End If
'    Next
    PrintMain 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim x%
'    sessionID_ = GetGUID
'    PopulateCompany
'    PopulateZone
'    InitWaste
'    dtFrom.DayValue = Format(Now, "dd")
    StatusBar1.Panels(1).Text = "......."
End Sub

Private Sub PrintMain(ByVal printType As Integer)
Dim Dates As String
Dim tSql As String
    Dates = Format(Date, "yyyy/mm/dd")
    With rpt1
        .Reset
        .WindowTitle = "รายงานใบเสนอราคา"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & uID & ";PWD = " & uPwd
        .ReportFileName = App.Path & "\Reports\QuotationRpt.rpt"
        .Formulas(0) = "RangDate = 'ช่วงวันที่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & "'"
        tSql = SqlQuery("Com_vwQuotationSearch", 0)
        tSql = tSql & " ORDER BY StaffFName, CompanyName, QuotationNo, WasteName"
        .SqlQuery = tSql
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

'Private Sub PopulateZone()
'Dim x%, tmpS As New BWGSearchManager
'    cboZone.Clear
'    cboZone.AddItem "ทุกโซน"
'    Set colZone_ = tmpS.SaleZoneSearch
'    For x = 1 To colZone_.Count
'        cboZone.AddItem colZone_(x).ZoneName
'    Next
'    If cboZone.ListCount > 0 Then
'        cboZone.ListIndex = 0
'    End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub

Private Sub Result_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Result.SortKey = ColumnHeader.Index - 1
    If Result.SortOrder = lvwAscending Then
        Result.SortOrder = lvwDescending
    Else
        Result.SortOrder = lvwAscending
    End If
    Result.Sorted = True
End Sub

Private Sub SettingResult(ByVal Index As Long, ByVal SKey As String, ByVal SText As String)
Dim i, j As Integer
    For i = 1 To Result.ListItems.Count
        If SKey = Result.ListItems.Item(i).Key Then
            Result.ListItems.Item(i).ForeColor = vbBlack
            For j = 1 To Result.ListItems.Item(i).ListSubItems.Count - 1
                Result.ListItems.Item(i).ListSubItems(j).ForeColor = vbBlack
            Next
            Exit For
        End If
    Next
End Sub

Private Sub txtCR_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtCR.Text = "" Then Exit Sub
    If txtCR.Text = "*" Then Exit Sub
    If KeyCode = vbKeyReturn Then
        txtCR.Text = GetStaffName(txtCR.Text, 2)
    End If
End Sub

Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCustomer_Validate False
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If txtCustomer.Text = "" Then Exit Sub
    If txtCustomer.Text = "*" Then Exit Sub
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    txtCustomer.Tag = ""
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frCustomerSearch
            With frCustomerSearch
                .getCriteria (txtCustomer.Text)
                .setCallerForm Me
                .Show 1
            End With
            If curCustomer_.ID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtSale.Text = "" Then Exit Sub
        'If txtSale.Text = "*" Then Exit Sub
        txtSale.Text = GetStaffName(txtSale.Text, 1)
    End If
End Sub

Private Function GetStaffName(ByVal StaffName As String, ByVal SType As Integer)
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If Trim(StaffName) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(StaffName) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(StaffName) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frStaffSearch
            With frStaffSearch
                .getCriteria StaffName, SType
                .setCallerForm Me
                .Show 1
            End With
            If curStaff_.ID = "" Then
                GetStaffName = ""
            Else
                GetStaffName = Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName)
            End If
        End If
    End If
End Function

Private Sub PopulateCompany()
Dim x%
    cboCompany.Clear
    cboCompany.AddItem "== เลือกทุกบริษัท =="
    For x = 1 To CurrentUser.WorkCompanies.Count
        cboCompany.AddItem CurrentUser.WorkCompanies(x).CompanyNameTH
    Next
    If cboCompany.ListCount > 0 Then
        cboCompany.ListIndex = 0
    End If
End Sub

'Private Sub InitWaste()
'Dim x%
'Dim strWaste As String
'Dim tWaste() As String
'    strWaste = "ไม่ระบุ,Haz,Non,HBL,NBL"
'    tWaste = Split(strWaste, ",")
'    cboWaste.Clear
'    For x = 0 To 4
'        cboWaste.AddItem tWaste(x)
'    Next
'    If cboWaste.ListCount > 0 Then
'        cboWaste.ListIndex = 0
'    End If
'End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    txtCustomer.Tag = ""
    With curCustomer_
        txtCustomer.Text = .CustomerName
        txtCustomer.Tag = .ID
    End With
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, ByVal StaffType As Integer, Optional ByVal isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curStaff_ = Nothing
    Set curStaff_ = selStaff
    If curStaff_ Is Nothing Then Exit Sub
    With curStaff_
        If StaffType = 1 Then
            txtSale.Text = .StaffFName & " " & .StaffLName
            SaleStaffID = curStaff_.ID
        ElseIf StaffType = 2 Then
            txtCR.Text = .StaffFName & " " & .StaffLName
            CRStaffID = curStaff_.ID
        End If
    End With
End Sub

Public Function SqlQuery(ByVal ComTable As String, Optional ByVal UseDate As Integer = 1) As String
Dim tmpSql As String, iCount As Long
    tmpSql = "Select Distinct * from " & ComTable & "" & vbCrLf
'    If UseDate = 1 Then
        tmpSql = tmpSql & "WHERE IssueDate >= '" & dtFrom.ValueYMD & "' And IssueDate <='" & dtTo.ValueYMD & "'" & vbCrLf
'    End If
    If cboCompany.ListIndex > 0 Then
        tmpSql = tmpSql & "AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex).ID & "'" & vbCrLf
    End If
    If Trim(txtCR.Text) <> "" Then
        If Trim(txtCR.Text) <> "*" Then
            tmpSql = tmpSql & " And IssuedStaffID ='" & CRStaffID & "'"
        End If
    End If
    If Trim(txtSale.Text) <> "" Then
        If Trim(txtSale.Text) <> "*" Then
            tmpSql = tmpSql & " And SaleStaffID ='" & SaleStaffID & "'"
        End If
    End If
    If Trim(txtCustomer.Text) <> "" Then
        If Trim(txtCustomer.Text) <> "*" Then
            tmpSql = tmpSql & " And CustomerID = '" & curCustomer_.ID & "'"
        End If
    End If
    If Trim(txtWaste.Text) <> "" Then
        If Trim(txtWaste.Text) <> "*" Then
            tmpSql = tmpSql & "AND WasteName LIKE '%" & Replace(txtWaste, "*", "") & "%'"
        End If
    End If
    tmpSql = tmpSql & vbCrLf
    SqlQuery = tmpSql
End Function

Private Sub SaveData()
On Error GoTo ErrD
Dim RS As ADODB.Recordset
Dim tSql As String
Dim CurrentDay As String
Dim i, j As Long
    CurrentDay = Format(Date, "yyyy/mm/dd")
    Counts = 0
    If Result.ListItems.Count <= 0 Then Exit Sub
    DBManager.GetDBConnection.BeginTrans
    For i = 1 To Result.ListItems.Count
        If Result.ListItems.Item(i).Checked = True Then
            Counts = Counts + 1
            tSql = "INSERT INTO tbCom_ExtCom_Paid(SDName , UnderSaleStaffID, StaffFName, CustomerID, CompanyName, "
            tSql = tSql & "WorkDate, MenifestID, DocNo, WasteDataID, WasteName, ReNo, REDate, Weight, TreatmentRate, "
            tSql = tSql & "TreatMentUnitID, InCollectFee, TransportFee, ComRateAmt, RateType, CommValue, VatValue, NetComm, NonOrHz, SaleZoneID, CompanyID, "
            tSql = tSql & " StaffLName,TimeTableNo,TreatmentCharge, WasteType, ZoneName, isCanceled, SaleDistrictID, ComRecName, ComRatePC, RateTypePC, DocCustWeight, PaidDate)" & vbCrLf
            tSql = tSql & "Values('"
            tSql = tSql & Result.ListItems.Item(i) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(1) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(2) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(3) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(4) & "', '"
            tSql = tSql & Format(Result.ListItems.Item(i).SubItems(5), "yyyy/mm/dd") & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(6) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(7) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(8) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(9) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(10) & "', '"
            tSql = tSql & Format(Result.ListItems.Item(i).SubItems(11), "yyyy/mm/dd") & "', "
            tSql = tSql & CDbl(Result.ListItems.Item(i).SubItems(12)) & ", "
            tSql = tSql & CDbl(Result.ListItems.Item(i).SubItems(13)) & ", '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(14) & "', "
            tSql = tSql & CDbl(Result.ListItems.Item(i).SubItems(15)) & ", "
            tSql = tSql & CDbl(Result.ListItems.Item(i).SubItems(16)) & ", "
            tSql = tSql & CDbl(Result.ListItems.Item(i).SubItems(17)) & ", '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(18) & "', "
            tSql = tSql & CDbl(Result.ListItems.Item(i).SubItems(19)) & ", "
            tSql = tSql & CDbl(Result.ListItems.Item(i).SubItems(20)) & ", "
            tSql = tSql & CDbl(Result.ListItems.Item(i).SubItems(21)) & ", '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(22) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(23) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(24) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(25) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(26) & "', "
            tSql = tSql & CDbl(Result.ListItems.Item(i).SubItems(27)) & ", '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(28) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(29) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(30) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(31) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(32) & "', "
            tSql = tSql & CDbl(Result.ListItems.Item(i).SubItems(33)) & ", '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(34) & "', '"
            tSql = tSql & Result.ListItems.Item(i).SubItems(35) & "', '"
            tSql = tSql & CurrentDay & "'"
            tSql = tSql & ")"
            DBManager.GetDBConnection.Execute tSql
        End If
    Next
    DBManager.GetDBConnection.CommitTrans
    MsgBox "บันทึกการตัดจ่ายค่าคอมฯ เสร็จแล้ว", vbInformation, "Completed"
    
    For i = 1 To Result.ListItems.Count
        If Result.ListItems.Item(i).Checked = True Then
            Result.ListItems.Item(i).Checked = False
            Result.ListItems.Item(i).ForeColor = vbBlue
            For j = 1 To Result.ColumnHeaders.Count - 1
                Result.ListItems.Item(i).ListSubItems(j).ForeColor = vbBlue
            Next
        End If
    Next
    Exit Sub
ErrD:
    MsgBox Err.Description
    DBManager.GetDBConnection.RollbackTrans
    Exit Sub
End Sub

Public Sub AddtoListView(ByRef RS As ADODB.Recordset, Optional ByVal UseDate As Byte = 1)
On Error GoTo ErrD
Dim tSql As String
Dim tList As ListItem
Dim i As Integer
   Result.Checkboxes = False
'    Set tList = Result.ListItems.Add(, SaleStaffID & "_" & CustomerID)
'    tList.SubItems(1) = SaleStaffID
'    tList.SubItems(2) = SaleName
'    tList.SubItems(3) = CustomerID
'    tList.SubItems(4) = CustomerName
'    tList.SubItems(5) = QCount
'    tList.SubItems(6) = Q2Count
        
    Do While Not RS.EOF
        StatusBar1.Panels(2).Text = "เพิ่มใบเสนอราคาเลขที่ " & RS!QuotationNo & " ลงในรายการ" & "  [" & RS.AbsolutePosition & "/" & RS.RecordCount & "]"
        Set tList = Result.ListItems.Add(, RS!IssuedStaffID & "_" & RS!CustomerID & "_" & RS!QuotationID & "_" & RS!WasteDataID, RS!StaffFName)
        tList.SubItems(1) = RS!QuotationNo
        tList.SubItems(2) = IIf(IsNull(RS!IssueDate), "", Format(RS!IssueDate, "dd/mm/yyyy"))
        tList.SubItems(3) = RS!CompanyName
        tList.SubItems(4) = RS!WasteName
        tList.SubItems(5) = IIf(IsNull(RS!WasteType), "", RS!WasteType)
        tList.SubItems(6) = IIf(IsNull(RS!TreatmentRate), "0", Format(RS!TreatmentRate, "#,##0"))
        tList.SubItems(7) = IIf(IsNull(RS!TreatmentUnitID), "0", RS!TreatmentUnitID)
        tList.SubItems(8) = IIf(IsNull(RS!TruckTypeDesc), "", RS!TruckTypeDesc)
        tList.SubItems(9) = IIf(IsNull(RS!TransportFee), "0", Format(RS!TransportFee, "#,##0"))
        tList.SubItems(10) = IIf(IsNull(RS!TrasnferUnitID), "", RS!TrasnferUnitID)
        tList.SubItems(11) = IIf(IsNull(RS!MinWeightPerCar), "0", RS!MinWeightPerCar)
        tList.SubItems(12) = IIf(IsNull(RS!StaffCountPerCar), "0,", RS!StaffCountPerCar)
        tList.SubItems(13) = IIf(IsNull(RS!ExtraLaborFee), "0", RS!ExtraLaborFee)
        tList.SubItems(14) = RS!QStatus
        
        If RS!QStatus = "ยกเลิก" Then
            tList.ForeColor = vbRed
            For i = 1 To Result.ColumnHeaders.Count - 1
                tList.ListSubItems(i).ForeColor = vbRed
            Next
        ElseIf RS!QStatus = "อนุมัติ" Then
            tList.ForeColor = vbBlue
            For i = 1 To Result.ColumnHeaders.Count - 1
                tList.ListSubItems(i).ForeColor = vbBlue
            Next
        End If
        
        RS.MoveNext
    Loop
    StatusBar1.Panels(2).Text = ""
    Set RS = Nothing
    Exit Sub
ErrD:
    MsgBox Err.Description
    Set RS = Nothing
    Exit Sub
End Sub

