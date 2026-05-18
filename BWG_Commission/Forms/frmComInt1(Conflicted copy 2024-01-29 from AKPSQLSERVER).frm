VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmComInt1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4095
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5055
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   915
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   60
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   0
      TabIndex        =   11
      Top             =   765
      Width           =   5955
      Begin VB.CheckBox ChkNewWaste 
         BackColor       =   &H8000000E&
         Caption         =   "Waste ใหม่"
         Height          =   195
         Left            =   1125
         TabIndex        =   30
         Top             =   3285
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Frame Frame2 
         Height          =   465
         Left            =   495
         TabIndex        =   21
         Top             =   2340
         Width           =   5145
         Begin VB.OptionButton Option4 
            Caption         =   "ทั้งหมด"
            Height          =   195
            Left            =   4005
            TabIndex        =   25
            Top             =   180
            Width           =   960
         End
         Begin VB.OptionButton Option3 
            Caption         =   "งาน DOC"
            Height          =   195
            Left            =   2835
            TabIndex        =   24
            Top             =   180
            Width           =   1365
         End
         Begin VB.OptionButton Option2 
            Caption         =   "งาน Job"
            Height          =   195
            Left            =   1620
            TabIndex        =   23
            Top             =   180
            Width           =   1365
         End
         Begin VB.OptionButton Option1 
            Caption         =   "งาน Routine"
            Height          =   195
            Left            =   180
            TabIndex        =   22
            Top             =   180
            Width           =   1365
         End
      End
      Begin VB.OptionButton OpOld 
         BackColor       =   &H80000009&
         Caption         =   "เขตการขายเก่า"
         Height          =   255
         Left            =   6825
         TabIndex        =   20
         Top             =   2385
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.OptionButton OpNow 
         BackColor       =   &H80000009&
         Caption         =   "เขตการขายปัจจุบัน"
         Height          =   255
         Left            =   5220
         TabIndex        =   19
         Top             =   1980
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Text            =   "*"
         Top             =   2040
         Width           =   4455
      End
      Begin VB.ComboBox cboWaste 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   4455
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   4
         Text            =   "*"
         Top             =   1320
         Width           =   4455
      End
      Begin VB.ComboBox cboZone 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   4455
      End
      Begin Commission.ctlDate dtFrom 
         Height          =   375
         Left            =   1140
         TabIndex        =   1
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin Commission.ctlDate dtTo 
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin VB.Frame Frame3 
         Height          =   465
         Left            =   495
         TabIndex        =   26
         Top             =   2745
         Width           =   5145
         Begin VB.OptionButton Option8 
            Caption         =   "งานลูกค้าใหม่"
            Height          =   195
            Left            =   180
            TabIndex        =   29
            Top             =   180
            Width           =   1365
         End
         Begin VB.OptionButton Option7 
            Caption         =   "งานลูกค้าเก่า"
            Height          =   195
            Left            =   2025
            TabIndex        =   28
            Top             =   180
            Width           =   1365
         End
         Begin VB.OptionButton Option5 
            Caption         =   "ทั้งหมด"
            Height          =   195
            Left            =   4005
            TabIndex        =   27
            Top             =   180
            Width           =   960
         End
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   18
         Top             =   2100
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ขน"
         Height          =   195
         Index           =   6
         Left            =   3240
         TabIndex        =   17
         Top             =   660
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ INV"
         Height          =   195
         Index           =   5
         Left            =   435
         TabIndex        =   16
         Top             =   660
         Width           =   660
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waste"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   15
         Top             =   1740
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   14
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โซน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   780
         TabIndex        =   13
         Top             =   1020
         Width           =   315
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   1380
         Width           =   735
      End
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
      Caption         =   "ค่าคอมภายใน"
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
      Left            =   150
      TabIndex        =   10
      Top             =   60
      Width           =   1185
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmComInt1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmComInt1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff
Dim colZone_ As Collection
Dim sessionID_ As String
Dim curCustomer_ As BWGCustomer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    PrintMain 0
End Sub

Private Sub cmdPrint_Click()
    PrintMain 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim x%
    sessionID_ = GetGUID
    PopulateCompany
    PopulateZone
    InitWaste
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    dtFrom.DayValue = "01"
    Option5.Value = True
    Option4.Value = True
End Sub

Private Sub PrintMain(printType As Integer)
Dim rs As New ADODB.Recordset
Dim tmpSql As String, iCount As Long
Dim tSql As String
    tmpSql = ""

    tmpSql = "SELECT     TOP (100) PERCENT SDName, StaffFName, StaffLName, CompanyName, TimeTableNo, WorkDate, DocNo, SUM(CustWeight) AS CustWeight," & vbCrLf
    tmpSql = tmpSql & "                      MAX(InCollectFee) AS InCollectFee, SUM(TreatmentCharge) AS TreatmentCharge, MAX(TransportFee) AS TransportFee, NonOrHz, WasteType," & vbCrLf
    tmpSql = tmpSql & "                      ZoneName , CompanyID, SaleZoneID, UnderSaleStaffID, CustomerID, isCanceled, SaleDistrictID, DocCustWeight, Sum(DisposerWeight)" & vbCrLf
    tmpSql = tmpSql & "                      AS DisposerWeight, FromAumphurID, FromProvinceID, FromAumphurName, FromProvinceName, ToAumphurID, ToProvinceID, ToAumphurName," & vbCrLf
    tmpSql = tmpSql & "                      ToProvinceName, TruckTypeID, TruckSubTypeID, MAX(ISNULL(CustRatePerTrip, 0)) AS CustRatePerTrip, MAX(ISNULL(TripTransFee_Inc, 0)) AS TripTransFee_Inc, MAX(ISNULL(VDTransFee, 0)) AS VDTransFee" & vbCrLf
    tmpSql = tmpSql & "From dbo.Com_vwComIn11" & vbCrLf
    If OpOld.Value = True Then
         tmpSql = tmpSql & "WHERE     (WorkDate >= '" & dtFrom.ValueYMD & "') And (WorkDate<='" & dtTo.ValueYMD & "') And  isActive <> 'Y'" & vbCrLf
    Else
        tmpSql = tmpSql & "WHERE     (WorkDate >= '" & dtFrom.ValueYMD & "') And (WorkDate<='" & dtTo.ValueYMD & "') And  isActive = 'Y'" & vbCrLf
    End If
    If cboZone.ListIndex > 0 Then
        tmpSql = tmpSql & " And SaleZoneID='" & colZone_(cboZone.ListIndex).ID & "'" & vbCrLf
    End If
    If Option2.Value = True Then
        tmpSql = tmpSql & " And  jobisjobs='Y'" & vbCrLf
    End If
    If Option1.Value = True Then
        tmpSql = tmpSql & " And jobisjobs='N'" & vbCrLf
    End If
    If Option3.Value = True Then
        tmpSql = tmpSql & " And jobisjobs='D'" & vbCrLf
    End If
    If Option4.Value = True Then
        tmpSql = tmpSql & " And (jobisjobs='D' or jobisjobs='N' or jobisjobs='Y')" & vbCrLf
    End If
    If Option8.Value = True Then
        tmpSql = tmpSql & " And (newcustjob='Y')" & vbCrLf
    End If
    If Option7.Value = True Then
        tmpSql = tmpSql & " And (newcustjob<>'Y')" & vbCrLf
    End If
    If Option5.Value = True Then
    
    End If
    If ChkNewWaste.Value = 1 Then
        tmpSql = tmpSql & " And (isnewwaste='New')" & vbCrLf
    End If
    If Trim(txtSale.Text) <> "" Then
        If Trim(txtSale.Text) <> "*" Then
            tmpSql = tmpSql & " And UnderSaleStaffID='" & curStaff_.ID & "'" & vbCrLf
        End If
    End If
    If cboWaste.ListIndex > 0 Then
        tmpSql = tmpSql & " And NonOrHz='" & Trim(cboWaste.Text) & "'" & vbCrLf
    End If
    If Trim(txtCustomer.Text) <> "" Then
        If Trim(txtCustomer.Text) <> "*" Then
            tmpSql = tmpSql & "And CustomerID = '" & txtCustomer.Tag & "'" & vbCrLf
        End If
    End If
    
    tSql = "SELECT * FROM Com_vwSales_District_Target" & vbCrLf
    tSql = tSql & "WHERE (CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "')" & vbCrLf
    If Not curStaff_ Is Nothing Then tSql = tSql & "AND SaleStaffID = '" & curStaff_.ID & "'" & vbCrLf
    tSql = tSql & "AND ForYear = " & dtFrom.YearValue & " AND ForMonth = " & dtFrom.MonthValue & vbCrLf
    
'    If CBS1.ListIndex = 0 Then
'        tSql = tSql & "AND ZoneName like 'AKP%' "
'    End If
'    If CBS1.ListIndex = 1 Then
'        tSql = tSql & "AND ZoneName like 'BME%' "
'    End If

    tSql = tSql & "ORDER BY SDName"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    
    If CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).CompanyCode = "AKP" Then
        tmpSql = tmpSql & "AND (CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "')" & vbCrLf
        If Not rs.EOF Then
            If rs!IsOverall = "0" Then tmpSql = tmpSql & "AND NonOrHz IN ('HIC', 'HBI', 'NIC', 'NBI', 'BBH', 'BBL','DOC','BWG','BLF')" & vbCrLf
        Else
            tmpSql = tmpSql & "AND NonOrHz IN ('HIC', 'HBI', 'NIC', 'NBI', 'BBH', 'BBL','DOC','BWG','BLF')" & vbCrLf
        End If
    Else
        tmpSql = tmpSql & "AND (CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + 1).ID & "')" & vbCrLf
        If Not rs.EOF Then
            If rs!IsOverall = "0" Then
                tmpSql = tmpSql & "AND NonOrHz NOT IN ('HBI', 'NIC', 'NBI', 'BBH', 'BBL','DOC','BWG','BLF')" & vbCrLf
            Else
                tmpSql = tmpSql & "AND NonOrHz IN ('HIC', 'HBI', 'NIC', 'NBI', 'BBH', 'BBL','DOC','BWG','BLF')" & vbCrLf
            End If
        Else
            tmpSql = tmpSql & "AND NonOrHz NOT IN ('HBI', 'NIC', 'NBI', 'BBH', 'BBL','DOC','BWG','BLF')" & vbCrLf
        End If
    End If
    
    '***Truck Type
'    tmpSql = tmpSql & "AND TruckSubTypeID IN (SELECT SubTypeID FROM tbTruckSubType WHERE TruckTypeID IN ( '67D1CCC9C8FF4782B89E3F56100A1592','D8B22F81A3D64E328407B736A27AEB3F' ))" & vbCrLf
    
    tmpSql = tmpSql & "GROUP BY SDName, StaffFName, StaffLName, CompanyName, TimeTableNo, WorkDate, DocNo, NonOrHz, WasteType, ZoneName, CompanyID, SaleZoneID," & vbCrLf
    tmpSql = tmpSql & "                      UnderSaleStaffID, CustomerID, isCanceled, SaleDistrictID, DocCustWeight, FromAumphurID, FromProvinceID, FromAumphurName, FromProvinceName," & vbCrLf
    tmpSql = tmpSql & "                      ToAumphurID , ToProvinceID, ToAumphurName, ToProvinceName, TruckTypeID, TruckSubTypeID, isPriceIncTransport" & vbCrLf
    tmpSql = tmpSql & "ORDER BY StaffLName, CompanyName, TimeTableNo, WorkDate, DocNo"
    
    With rpt1
        .Reset
        .WindowTitle = "รายงานการตั้งค่าคอมฯใน"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .WindowShowGroupTree = True
        .WindowShowPrintBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If dtTo.YearValue < "2564" Then
        .ReportFileName = App.Path & "\Reports\ComIn_Old.rpt"
        Else
        .ReportFileName = App.Path & "\Reports\ComIn.rpt"
        End If
        .Formulas(0) = "Company='บริษัท " & Trim(cboCompany.Text) & "'"
        .Formulas(1) = "WorkDate='" & "วันที่  " & dtFrom.ValueDMY & "  ถึงวันที่  " & dtTo.ValueDMY & "'"
        .Formulas(2) = "Sales='" & IIf(Trim(txtSale.Text) <> "", IIf(Trim(txtSale) <> "*", txtSale.Text, ""), "") & "'"
        .SqlQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
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

Private Sub PopulateZone()
Dim x%, tmpS As New BWGSearchManager
    cboZone.Clear
    cboZone.AddItem "ทุกโซน"
    Set colZone_ = tmpS.SaleZoneSearch
    For x = 1 To colZone_.Count
        cboZone.AddItem colZone_(x).ZoneName
    Next
    If cboZone.ListCount > 0 Then
        cboZone.ListIndex = 0
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DBConnExc "Delete tbPrintOrd01 Where SessionID='" & sessionID_ & "'"
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

Private Sub txtSale_Validate(Cancel As Boolean)
    If txtSale.Text = "" Then Exit Sub
    If txtSale.Text = "*" Then Exit Sub
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSale.Text) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frStaffSearch
            With frStaffSearch
                .getCriteria (txtSale.Text)
                .setCallerForm Me
                .Show 1
            End With
            If curStaff_.ID = "" Then txtSale.Text = ""
        End If
    End If
End Sub

Private Sub PopulateCompany()
Dim x%
    cboCompany.Clear
    For x = 1 To CurrentUser.WorkCompanies.Count
        cboCompany.AddItem CurrentUser.WorkCompanies(x).CompanyNameTH
    Next
    If cboCompany.ListCount > 0 Then
        cboCompany.ListIndex = 0
    End If
End Sub

Private Sub InitWaste()
Dim x%
Dim strWaste As String
Dim tWaste() As String
    strWaste = "ไม่ระบุ,Haz,Non,HBL,NBL"
    tWaste = Split(strWaste, ",")
    cboWaste.Clear
    For x = 0 To 4
        cboWaste.AddItem tWaste(x)
    Next
    If cboWaste.ListCount > 0 Then
        cboWaste.ListIndex = 0
    End If
End Sub

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

