VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmPrintAllCustomerPrice 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "รายงานลูกค้าใช้บริการ"
   ClientHeight    =   3960
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport rpt1 
      Left            =   45
      Top             =   3375
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   45
      TabIndex        =   4
      Top             =   855
      Width           =   6720
      Begin VB.ComboBox cboYear 
         Height          =   315
         Left            =   1170
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   225
         Width           =   1455
      End
      Begin VB.ComboBox cboZone 
         Height          =   315
         Left            =   1170
         TabIndex        =   8
         Text            =   "Combo2"
         Top             =   720
         Width           =   1230
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1185
         TabIndex        =   7
         Top             =   1215
         Width           =   5415
      End
      Begin VB.ComboBox cboCustType 
         Height          =   315
         ItemData        =   "frmPrintAllCustomerPrice.frx":0000
         Left            =   1155
         List            =   "frmPrintAllCustomerPrice.frx":0002
         TabIndex        =   6
         Top             =   1755
         Width           =   1815
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3600
         TabIndex        =   5
         Top             =   720
         Width           =   3000
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   630
         TabIndex        =   14
         Top             =   1260
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลือกปี :"
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   13
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โซนการขาย :"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   765
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย :"
         Height          =   195
         Index           =   3
         Left            =   2565
         TabIndex        =   11
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทลูกค้า :"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   10
         Top             =   1755
         Width           =   1020
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   4830
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3105
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5790
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3105
      Width           =   915
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   795
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3105
      Width           =   915
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
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานลูกค้าใช้บริการ"
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
      TabIndex        =   0
      Top             =   60
      Width           =   1890
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
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frmPrintAllCustomerPrice.frx":0004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9150
   End
End
Attribute VB_Name = "frmPrintAllCustomerPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colZone_ As Collection
Dim curStaff_ As BWGCompanyStaff
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
    Dim selYears As Integer
    Dim i
    With cboYear
        selYears = Year(Date) - 3
        If selYears < 2500 Then selYears = selYears + 543
        For i = 1 To 4
        selYears = selYears + 1
        .AddItem selYears
        Next
        .ListIndex = 2
    End With

    With cboCustType
        .AddItem "== ประเภทลูกค้า =="
        .AddItem "(C) ,(s)"
        .AddItem "(M) ,(MS)"
        .AddItem "(B) ,(E)"
        .ListIndex = 0
    End With
PopulateZone
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

Private Sub PrintMain(printType As Integer)
Dim tmpSql As String
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\ReportMTAll.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_AllCustomerPrice " & vbCrLf
        tmpSql = tmpSql & "WHERE WorkYear='" & cboYear.Text & "' "
                
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
         If cboCustType.ListIndex = 1 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('C','S') "
        ElseIf cboCustType.ListIndex = 2 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('M','MS')"
        ElseIf cboCustType.ListIndex = 3 Then
            tmpSql = tmpSql & " AND CustTypeFlag in ('B','E')"
        End If
        
        tmpSql = tmpSql & vbCrLf & "ORDER BY SaleStaffFName,CompanyName"
'        .Formulas(0) = "FromDate=" & Chr(34) & "ข้อมูลวันที่ " & " " & dtFrom.ValueDMY & " " & " ถึงวันที่ " & " " & dtTo.ValueDMY & Chr(34)
'        .Formulas(0) = "PrintedDate=" & Chr(34) & ข้อมูลวันที่ & FormatYMD_to_DMY(TodayDate, "/", "/") & Chr(34)
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

