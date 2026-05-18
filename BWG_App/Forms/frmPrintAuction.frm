VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmPrintAuction 
   Caption         =   "งานประมูล"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "การเรียงลำดับของข้อมูล"
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   90
      TabIndex        =   17
      Top             =   3510
      Visible         =   0   'False
      Width           =   5880
      Begin VB.ComboBox cboSort1 
         Height          =   315
         ItemData        =   "frmPrintAuction.frx":0000
         Left            =   540
         List            =   "frmPrintAuction.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   360
         Width           =   2385
      End
      Begin VB.ComboBox cboSort2 
         Height          =   315
         ItemData        =   "frmPrintAuction.frx":0051
         Left            =   3300
         List            =   "frmPrintAuction.frx":0061
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   360
         Width           =   2385
      End
      Begin VB.ComboBox cboSort3 
         Height          =   315
         ItemData        =   "frmPrintAuction.frx":00A2
         Left            =   540
         List            =   "frmPrintAuction.frx":00B2
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   750
         Width           =   2385
      End
      Begin VB.ComboBox cboSort4 
         Height          =   315
         ItemData        =   "frmPrintAuction.frx":00F3
         Left            =   3300
         List            =   "frmPrintAuction.frx":0103
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   750
         Width           =   2385
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1."
         Height          =   195
         Index           =   4
         Left            =   330
         TabIndex        =   25
         Top             =   420
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2."
         Height          =   195
         Index           =   5
         Left            =   3090
         TabIndex        =   24
         Top             =   420
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3."
         Height          =   195
         Index           =   6
         Left            =   330
         TabIndex        =   23
         Top             =   810
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4."
         Height          =   195
         Index           =   7
         Left            =   3090
         TabIndex        =   22
         Top             =   810
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3150
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2670
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2670
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4110
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2670
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1845
      Left            =   60
      TabIndex        =   0
      Top             =   810
      Width           =   5925
      Begin VB.ComboBox cboDateType 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2250
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1110
         TabIndex        =   4
         Top             =   1035
         Width           =   4425
      End
      Begin VB.ComboBox cboZone 
         Height          =   315
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   645
         Width           =   4455
      End
      Begin VB.CheckBox CheckNoMenifest 
         BackColor       =   &H80000005&
         Caption         =   "มีใบอนุญาติแล้วแต่ยังไม่มีการขน"
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   1410
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.ComboBox CboCompany 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Text            =   "Cbocompany"
         Top             =   1920
         Visible         =   0   'False
         Width           =   4455
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1425
         TabIndex        =   6
         Top             =   255
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3780
         TabIndex        =   7
         Top             =   255
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3225
         TabIndex        =   13
         Top             =   315
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่"
         Height          =   195
         Index           =   3
         Left            =   450
         TabIndex        =   12
         Top             =   2310
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ประมูลตั้งแต่"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   11
         Top             =   285
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โซน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   720
         TabIndex        =   10
         Top             =   705
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   9
         Top             =   1095
         Width           =   720
      End
      Begin VB.Label Label3 
         Caption         =   "บริษัท :"
         Height          =   300
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   60
      Top             =   2850
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
      Caption         =   "รายงานการประมูล"
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
      TabIndex        =   26
      Top             =   60
      Width           =   1590
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
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmPrintAuction.frx":0144
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintAuction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim colZone_ As Collection
Dim curSale_ As BWGCompanyStaff
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
'    cboSort1.ListIndex = 0
'    cboSort2.ListIndex = 1
'    cboSort3.ListIndex = 2
'    cboSort4.ListIndex = 3
    With cboDateType
        .AddItem "ประมูล"
        .AddItem "คีย์ข้อมูล"
        .ListIndex = 0
    End With
    PopulateZone
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    PopulateCompany CboCompany, cboIndex
End Sub

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
Dim CusName As String
Dim sortField(0 To 3) As String
'    sortField(0) = "CompanyName"
'    sortField(1) = "StaffFName,StaffLName"
'    sortField(2) = "StartDate"
'    sortField(3) = "EndDate"
'    If cboSort1.ListIndex = cboSort2.ListIndex Or cboSort1.ListIndex = cboSort3.ListIndex Or cboSort1.ListIndex = cboSort4.ListIndex _
'        Or cboSort2.ListIndex = cboSort3.ListIndex Or cboSort2.ListIndex = cboSort4.ListIndex Or cboSort3.ListIndex = cboSort4.ListIndex Then
'            MsgBox "การเรียงลำดับข้อมูลในแต่ละตำแหน่งจะต้องไม่ซ้ำกัน", vbExclamation
'            cboSort1.SetFocus
'            Exit Sub
'    End If
'    If CboCompany.ListCount > 1 Then
'        If CboCompany.ListIndex > 0 Then
'                CusName = " And CompanyID='" & CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).ID & "' "
'        End If
'    Else
'        If CboCompany.ListIndex >= 0 Then
'                CusName = " And CompanyID='" & CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).ID & "' "
'        End If
'    End If
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\auction.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_PrintAuction " & vbCrLf
        If cboDateType.ListIndex = 0 Then
            tmpSql = tmpSql & "WHERE StartDate>='" & dtFrom.ValueYMD & "' And StartDate<='" & dtTo.ValueYMD & "' "
        ElseIf cboDateType.ListIndex = 1 Then
            tmpSql = tmpSql & "WHERE EndDate>='" & dtFrom.ValueYMD & "' And EndDate<='" & dtTo.ValueYMD & "' "
        End If
        If cboZone.ListIndex > 0 Then
            tmpSql = tmpSql & "And SaleZoneID='" & colZone_(cboZone.ListIndex).ID & "' "
        End If
        If Not curSale_ Is Nothing Then
            If Trim(curSale_.ID) <> "" Then
                tmpSql = tmpSql & "And SaleStaffID='" & curSale_.ID & "' "
            End If
        End If
        If CheckNoMenifest.Value = 1 Then
            tmpSql = tmpSql & "And (CustomerID Not in (Select CustomerID From tbMenifestHeader) )"
        End If
'        tmpSql = tmpSql & CusName
'        tmpSql = tmpSql & vbCrLf & " ORDER BY " & sortField(cboSort1.ListIndex) & "," & sortField(cboSort2.ListIndex) & "," & sortField(cboSort3.ListIndex) & "," & sortField(cboSort4.ListIndex)
        .Formulas(0) = "DateRange=" & Chr(34) & "ในช่วงวันที่" & cboDateType.Text & "ตั้งแต่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .Formulas(2) = "ShowZone=" & Chr(34) & cboZone.Text & Chr(34)
        If Not curSale_ Is Nothing Then
            If Trim(curSale_.ID) <> "" Then
                .Formulas(3) = "ShowSale=" & Chr(34) & "เฉพาะ " & txtSale.Text & Chr(34)
            Else
                .Formulas(3) = "ShowSale=" & Chr(34) & "" & Chr(34)
            End If
        Else
            .Formulas(3) = "ShowSale=" & Chr(34) & "" & Chr(34)
        End If
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub


'Private Sub txtCustomer_Validate(Cancel As Boolean)
'    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
'    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
'        Set curCustomer_ = Nothing
'        Set curCustomer_ = New BWGCustomer
'    Else
'        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
'            Set curCustomer_ = Nothing
'            Set curCustomer_ = New BWGCustomer
'            Load frCustomerSearch
'            With frCustomerSearch
'                .CStatus = "N"
'                .getCriteria (txtCustomer.Text)
'                .SetCallerForm Me
'                .Show 1
'            End With
'            If curCustomer_.ID = "" Then txtCustomer.Text = ""
'        End If
'    End If
'End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
Dim selStr As String
    If curSale_ Is Nothing Then Set curSale_ = New BWGCompanyStaff
    selStr = Trim(curSale_.StaffFName & " " & curSale_.StaffLName)
    If Trim(txtSale.Text) = "" And Trim(selStr) <> Trim(txtSale.Text) Then
        Set curSale_ = Nothing
        Set curSale_ = New BWGCompanyStaff
    Else
        If Trim(selStr) <> Trim(txtSale.Text) Then
            Set curSale_ = Nothing
            Set curSale_ = New BWGCompanyStaff
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtSale.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curSale_.ID = "" Then txtSale.Text = ""
        End If
    End If
End Sub
Private Sub PopulateZone()
Dim x%, tmpS As New BWGSearchManager
    Set colZone_ = tmpS.SaleZoneSearch
    cboZone.Clear
    cboZone.AddItem "ทุก Zone"
    For x = 1 To colZone_.Count
        cboZone.AddItem colZone_(x).ZoneName
    Next
    cboZone.ListIndex = 0
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff)
    Set curSale_ = selStaff
    With curSale_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub



