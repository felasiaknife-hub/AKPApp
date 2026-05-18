VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmPrintTRN03A 
   Caption         =   "TRN03A"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4755
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4755
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4755
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "รายงานการขนส่งรายวัน"
      ForeColor       =   &H80000008&
      Height          =   3930
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   5925
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000E&
         Caption         =   "รายงาน Forecast"
         Height          =   195
         Left            =   3285
         TabIndex        =   31
         Top             =   3600
         Width           =   2310
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000E&
         Height          =   465
         Left            =   270
         TabIndex        =   26
         Top             =   585
         Width           =   5415
         Begin VB.OptionButton Option4 
            BackColor       =   &H8000000E&
            Caption         =   "DOC"
            Height          =   195
            Left            =   4275
            TabIndex        =   30
            Top             =   180
            Width           =   870
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H8000000E&
            Caption         =   "JOB"
            Height          =   195
            Left            =   2745
            TabIndex        =   29
            Top             =   180
            Width           =   870
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H8000000E&
            Caption         =   "Routine"
            Height          =   195
            Left            =   1440
            TabIndex        =   28
            Top             =   180
            Width           =   870
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H8000000E&
            Caption         =   "ทั้งหมด"
            Height          =   195
            Left            =   135
            TabIndex        =   27
            Top             =   180
            Width           =   870
         End
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmPrintTRN03A.frx":0000
         Left            =   1140
         List            =   "frmPrintTRN03A.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1410
         Width           =   4395
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1050
         Width           =   4395
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   6
         Top             =   2490
         Width           =   4395
      End
      Begin VB.ComboBox cboZone 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2130
         Width           =   4395
      End
      Begin VB.ComboBox cboCarType 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1770
         Width           =   4395
      End
      Begin VB.TextBox txtVendor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   3
         Top             =   2850
         Width           =   4395
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   3225
         Width           =   4395
      End
      Begin VB.CheckBox ChkRdf 
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   3585
         Width           =   255
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1140
         TabIndex        =   9
         Top             =   270
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   10
         Top             =   270
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่ขน"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   330
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   19
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "บริษัท"
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   18
         Top             =   1110
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทกาก"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   17
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   16
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โซน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   720
         TabIndex        =   15
         Top             =   2190
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทรถ"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   14
         Top             =   1830
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor"
         Height          =   195
         Index           =   6
         Left            =   480
         TabIndex        =   13
         Top             =   2910
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   12
         Top             =   3285
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "แสดงเฉพาะ RDF"
         Height          =   195
         Index           =   9
         Left            =   1560
         TabIndex        =   11
         Top             =   3585
         Width           =   1230
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานตารางการขนส่งรายวัน (TRN03A)"
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
      TabIndex        =   25
      Top             =   60
      Width           =   3600
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   120
      X2              =   5760
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทรถ"
      Height          =   195
      Index           =   8
      Left            =   1560
      TabIndex        =   24
      Top             =   4200
      Width           =   795
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmPrintTRN03A.frx":0004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintTRN03A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colZone_ As Collection
Dim coltruckType_ As Collection
Dim curStaff_ As BWGCompanyStaff
Dim curVendor_ As BWGTransportVendor
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

Private Sub Form_Load()
Dim x As Integer
    PopulateCompany cboCompany, cboIndex
    PopulateCarType
    PopulateZone
    
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    Option1.Value = True
    With cboType
        .AddItem "==ทุกประเภท=="
        For x = 1 To WasteTypes.Count
            .AddItem WasteTypes(x).WasteType
        Next
        .ListIndex = 0
    End With
    
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        If Check1.Value = 1 Then
        .ReportFileName = App.Path & "\reports\TRN03A2.rpt"
        Else
        .ReportFileName = App.Path & "\reports\TRN03A.rpt"
        End If
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_RptTRN03A " & vbCrLf
        tmpSql = tmpSql & "WHERE WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "' "
        
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                    tmpSql = tmpSql & " AND JobNo NOT LIKE '%PST%'"
                ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                    tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                Else
                    tmpSql = tmpSql & "And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                    tmpSql = tmpSql & " AND JobNo NOT LIKE '%PST%'"
                End If
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    tmpSql = tmpSql & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                    tmpSql = tmpSql & " AND JobNo NOT LIKE '%PST%'"
                ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                    tmpSql = tmpSql & "And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                Else
                    tmpSql = tmpSql & "And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' "
                    tmpSql = tmpSql & " AND JobNo NOT LIKE '%PST%'"
                End If
            End If
        End If
        
        If cboType.ListIndex > 0 Then
            tmpSql = tmpSql & " AND WasteType='" & cboType.Text & "' "
        End If
        If Not curCustomer_ Is Nothing Then
            If Trim(curCustomer_.ID) <> "" Then
                tmpSql = tmpSql & " And CustomerID='" & curCustomer_.ID & "' "
            End If
        End If
        If cboZone.ListIndex > 0 Then
            tmpSql = tmpSql & " AND SaleZoneID='" & colZone_(cboZone.ListIndex).ID & "'"
        End If
        If Trim(txtSale.Text) <> "" Then
            tmpSql = tmpSql & " AND UnderSaleStaffID='" & curStaff_.ID & "'"
        End If
        
        If cboCarType.ListIndex > 0 Then
            tmpSql = tmpSql & " AND TruckTypeDesc = '" & Trim(cboCarType.Text) & "' "
        End If
               
        If Trim(txtVendor.Text) = "" Then Set curVendor_ = Nothing
        If Not curVendor_ Is Nothing Then
            If Trim(curVendor_.ID) <> "" Then
                tmpSql = tmpSql & " And VendorID='" & Trim(curVendor_.ID) & "' "
            End If
        End If
        If ChkRdf.Value = 1 Then
            tmpSql = tmpSql & " And treatmentDesc like '%RDF%' "
        End If
        Dim Job As String
        If Option2.Value = True Then
        tmpSql = tmpSql & " and isjobs='N' "
        ElseIf Option3.Value = True Then
        tmpSql = tmpSql & " and isjobs='Y' "
        ElseIf Option4.Value = True Then
        tmpSql = tmpSql & " and isjobs='D' "
        End If

        tmpSql = tmpSql & vbCrLf & "ORDER BY WorkDate,CompanyName,CarRegisID"
        .Formulas(0) = "FromDate=" & Chr(34) & "ข้อมูลวันที่ " & " " & dtFrom.ValueDMY & " " & " ถึงวันที่ " & " " & dtTo.ValueDMY & Chr(34)
'        .Formulas(0) = "PrintedDate=" & Chr(34) & ข้อมูลวันที่ & FormatYMD_to_DMY(TodayDate, "/", "/") & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
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
Dim tmpWaste As BWGWasteDataCR
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

Private Sub PopulateCarType()
Dim x%, tmpS As New BWGSearchManager
    cboCarType.Clear
    Set coltruckType_ = tmpS.TruckTypeSearch
    cboCarType.AddItem "==ระบุประเภทรถ=="
    For x = 1 To coltruckType_.Count
        cboCarType.AddItem coltruckType_(x).TypeDesc
    Next
    If cboCarType.ListCount > 0 Then
        cboCarType.ListIndex = 0
    End If
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


