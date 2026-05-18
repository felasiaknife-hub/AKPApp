VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmPrintFIN11 
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4350
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   5985
      Begin VB.ComboBox cboZone 
         Height          =   315
         Left            =   1140
         TabIndex        =   18
         Top             =   1800
         Width           =   4455
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   15
         Top             =   2280
         Width           =   4425
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   13
         Top             =   1320
         Width           =   4425
      End
      Begin VB.ComboBox cboCustType 
         Height          =   315
         Left            =   1140
         TabIndex        =   12
         Top             =   2760
         Width           =   1815
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   900
         Width           =   4425
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   1170
         TabIndex        =   2
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   3810
         TabIndex        =   3
         Top             =   360
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โซน"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   17
         Top             =   1800
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลุกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Top             =   2340
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้แทนขาย"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   1380
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทลูกค้า"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ช่วงวันที่ขน"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   6
         Top             =   420
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   5
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ขนโดย"
         Height          =   195
         Index           =   4
         Left            =   570
         TabIndex        =   4
         Top             =   960
         Width           =   495
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
      Caption         =   "พิมพ์รายงานเมนิเฟส ที่วางบิล แต่ยังไม่ได้รับเงิน (RV)"
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
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   5700
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   5640
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmPrintFIN11.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11280
   End
End
Attribute VB_Name = "frmPrintFIN11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim coltruckType_ As Collection
Dim curStaff_ As BWGCompanyStaff
Dim curCustomer_ As BWGCustomer
Dim colZone_ As Collection

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub TxtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then txtCustomer_Validate False
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
   If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(TxtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(TxtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(TxtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .getCriteria (TxtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_.ID = "" Then TxtCustomer.Text = ""
        End If
    End If
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
Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With Rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\FIN11.rpt"
        
        tmpSql = "Select * From vw_Me_Bill_RE " & vbCrLf
        tmpSql = tmpSql & " Where (WorkDate between '" & dtFrom.ValueYMD & "' and '" & dtTo.ValueYMD & "') " & vbCrLf
        tmpSql = tmpSql & " And isCanceled <> 'Y'   And INVNO is null  and SaleDistrictID <> '' " & vbCrLf
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                tmpSql = tmpSql & "AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                tmpSql = tmpSql & "AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'" & vbCrLf
            End If
        End If
    If cboZone.ListIndex > 0 Then
        tmpSql = tmpSql & " And ZoneName='" & cboZone.Text & "'" & vbCrLf
    End If
         If Trim(txtSale.Text) <> "" Then tmpSql = tmpSql & " AND UnderSaleStaffID='" & curStaff_.ID & "' " & vbCrLf
         If Not curCustomer_ Is Nothing Then
            If Trim(curCustomer_.ID) <> "" Then
                tmpSql = tmpSql & " And CustomerID='" & curCustomer_.ID & "' "
            End If
        End If
        If cboCustType.ListIndex = 1 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('C','S') "
        ElseIf cboCustType.ListIndex = 2 Then
            tmpSql = tmpSql & " AND CustTypeFlag = 'F' "
        ElseIf cboCustType.ListIndex = 3 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('L','P') "
        ElseIf cboCustType.ListIndex = 4 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('T','K') "
        ElseIf cboCustType.ListIndex = 5 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('G','E') "
        ElseIf cboCustType.ListIndex = 6 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('R') "
        End If
        tmpSql = tmpSql & " Order by WorkDate,TimetableNo,CustomerName "
        .Formulas(0) = "Company=" & Chr(34) & cboCompany.Text & Chr(34)
        .Formulas(1) = "DateReport=" & Chr(34) & "ในช่วงวันที่ขนตั้งแต่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34)
        .WindowTitle = "รายงานเมนิเฟสที่วางบิล แต่ยังไม่ได้รับเงิน (FIN11) "
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub cmdPreview_Click()
PrintMain 0
End Sub

Private Sub cmdPrint_Click()
 PrintMain 0
End Sub

Private Sub Form_Load()
    PopulateCompany cboCompany, cboIndex
    PopulateZone
     With frmPicture
        cmdPreview.Picture = .picPreview.Picture
        cmdPrint.Picture = .picPrint.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    dtFrom.DayValue = 1
    With cboCustType
        .AddItem "== ประเภทลูกค้า =="
        .AddItem "(C) ,(s)"
        .AddItem "(F) RDF Waste"
        .AddItem "(L) (p)"
        .AddItem "(T) (k)"
        .AddItem "(G) (e)"
        .AddItem "(R)"
        .ListIndex = 0
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
Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
Dim tmpWaste As BWGWasteDataCR
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    With curCustomer_
        TxtCustomer.Text = .CustomerName
    End With
End Sub
