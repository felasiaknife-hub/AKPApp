VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmCustomerReport01 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "รายชื่อลูกค้า"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check5 
      Caption         =   "แสดงยอดขาย"
      Height          =   255
      Left            =   180
      TabIndex        =   15
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "พนักงานขาย"
      Height          =   675
      Left            =   2520
      TabIndex        =   14
      Top             =   1860
      Width           =   4935
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Text            =   "*"
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "บริษัทที่รับกำจัด"
      Height          =   675
      Left            =   2520
      TabIndex        =   13
      Top             =   1200
      Width           =   4935
      Begin VB.ComboBox cboCompany 
         Height          =   315
         ItemData        =   "frmCustomerReport01.frx":0000
         Left            =   120
         List            =   "frmCustomerReport01.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   4665
      End
   End
   Begin Crystal.CrystalReport Rpt1 
      Left            =   60
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CheckBox Check4 
      Caption         =   "เลือกพนักงานขาย"
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   1980
      Width           =   2175
   End
   Begin VB.CheckBox Check3 
      Caption         =   "เลือกบริษัทที่รับกำจัด"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ออก"
      Height          =   375
      Left            =   6300
      TabIndex        =   10
      Top             =   3180
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   375
      Left            =   4950
      TabIndex        =   9
      Top             =   3180
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   2520
      TabIndex        =   1
      Top             =   180
      Width           =   4935
      Begin VB.CheckBox Check2 
         Caption         =   "แยกตามปีที่ขน"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   660
         Width           =   1395
      End
      Begin WasteManagment.ctlDate ctlFrom 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteManagment.ctlDate ctlTo 
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         Caption         =   "ถึง"
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   12
         Top             =   300
         Width           =   195
      End
      Begin VB.Label Label1 
         Caption         =   "วันที่ขน"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "เลือกเฉพาะลูกค้าที่มีการขน"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.Line Line1 
      X1              =   -60
      X2              =   7815
      Y1              =   3000
      Y2              =   3000
   End
End
Attribute VB_Name = "frmCustomerReport01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curStaff_ As BWGCompanyStaff

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        Frame1.Visible = False
    ElseIf Check1.Value = 1 Then
        Frame1.Visible = True
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 0 Then
        Frame2.Visible = False
    ElseIf Check3.Value = 1 Then
        Frame2.Visible = True
    End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = 0 Then
        Frame3.Visible = False
    ElseIf Check4.Value = 1 Then
        Frame3.Visible = True
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    Dim CountYY As Integer, i As Integer
    Dim FrYYMMDD As String, ToYYMMDD As String
    If Check1.Value = 1 Then
        If ctlTo.YearValue > ctlFrom.YearValue Then
            CountYY = ctlTo.YearValue - ctlFrom.YearValue + 1
        Else
            CountYY = 1
        End If
    Else
        Call PrintRpt
    End If
    
    If Check2.Value = 1 Then
        For i = 0 To CountYY - 1
            If ctlFrom.YearValue + i = ctlFrom.YearValue Then
                FrYYMMDD = ctlFrom.ValueYMD
            Else
                FrYYMMDD = ctlFrom.YearValue + i & "/" & "01" & "/" & "01"
            End If
            
            If ctlFrom.YearValue + i = ctlTo.YearValue Then
                ToYYMMDD = ctlFrom.YearValue + i & "/" & Format(ctlTo.MonthValue, "00") & "/" & Format(ctlTo.DayValue, "00")
            Else
                ToYYMMDD = ctlFrom.YearValue + i & "/" & "12" & "/" & "31"
            End If
            Call PrintRpt(FrYYMMDD, ToYYMMDD)
        Next i
    Else
        Call PrintRpt(ctlFrom.ValueYMD, ctlTo.ValueYMD)
    End If
    
End Sub

Private Sub PrintRpt(Optional ByVal FrYYMMDD As String = "0001/01/01", Optional ByVal ToYYMMDD As String = "9999/12/31")
    Dim RS As New ADODB.Recordset
    Dim tSql As String, strCond As String, CompanyID As String, SaleID As String
    Dim ShowIncome As String
    CompanyID = "": SaleID = ""
    tSql = "SELECT DISTINCT     CustomerCode, CompanyName, Addr1, TumbolName, AumphurName, ProvinceName, ZipCode, TelNo, FaxNo, ContactPerson, ContactDeptNo, EmailAddr, SaleName" & vbCrLf
    If Check5.Value = 1 Then
        tSql = tSql & ", Income = (SELECT SUM(IncollectFee) FROM dbo.tbMenifestHeader" & vbCrLf
        tSql = tSql & "                            WHERE CustomerID = vw_CustomerAddress.CustomerID AND ISNULL(isCanceled,'N') <> 'Y' AND isClosed = 'Y'" & vbCrLf
        tSql = tSql & "AND WorkDate >= '" & FrYYMMDD & "' AND WorkDate <= '" & ToYYMMDD & "')" & vbCrLf
    End If
    tSql = tSql & "From vw_CustomerAddress" & vbCrLf
    tSql = tSql & "WHERE     (CustomerID IN (SELECT CustomerID FROM dbo.tbMenifestHeader" & vbCrLf
    tSql = tSql & "                            WHERE ISNULL(isCanceled,'N') <> 'Y' AND isClosed = 'Y'" & vbCrLf
    tSql = tSql & "AND WorkDate >= '" & FrYYMMDD & "' AND WorkDate <= '" & ToYYMMDD & "'" & vbCrLf
    If Frame2.Visible = True Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    strCond = " AND (NonOrHz IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI') OR CompanyID = '776EDA5830F744A6B3551348D851BC22')"
                    CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
                Else
                    strCond = " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'"
                    strCond = strCond & " AND NonOrHz NOT IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')"
                    CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
                End If
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    strCond = " AND (NonOrHz IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI') OR CompanyID = '776EDA5830F744A6B3551348D851BC22')"
                    CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
                Else
                    strCond = " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'"
                    strCond = strCond & " AND NonOrHz NOT IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')"
                    CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
                End If
            End If
        End If
    End If
    tSql = tSql & strCond & "))" & vbCrLf
    If Frame3.Visible = True Then
        If Not curStaff_ Is Nothing Then
            If curStaff_.ID <> "" Then
                tSql = tSql & "AND UnderSaleStaffID = '" & curStaff_.ID & "' "
                SaleID = curStaff_.ID
            End If
        End If
    End If
    tSql = tSql & "AND (Status = 'N' OR Status = 'P')" & vbCrLf
    tSql = tSql & "ORDER BY CompanyName"
'    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
'    If RS.RecordCount > 0 Then
    If Check5.Value = vbChecked Then
        ShowIncome = "Y"
    Else
        ShowIncome = "N"
    End If
    tSql = "EXEC SP_CustomerAddressIncomes '" & FrYYMMDD & "', '" & ToYYMMDD & "', '" & "" & CompanyID & "', '" & SaleID & "', '" & ShowIncome & "'"
'    DBConn.Execute tSql
        With rpt1
            .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
            .ReportFileName = App.Path & "\Reports\CRList02.rpt"
            .ReportTitle = Me.Caption & "ที่มีการขนตั้งแต่ " & FrYYMMDD & " - " & ToYYMMDD
            .ParameterFields(0) = "@SDate; " & FrYYMMDD & "; True"
            .ParameterFields(1) = "@EDate; " & ToYYMMDD & "; True"
            .ParameterFields(2) = "@CompanyID; " & CompanyID & "; True"
            .ParameterFields(3) = "@SaleID; " & SaleID & "; True"
            .ParameterFields(4) = "@ShowIncome; " & ShowIncome & "; True"
            .WindowShowCloseBtn = True
            .WindowShowExportBtn = True
            .WindowShowPrintBtn = True
            .WindowShowPrintSetupBtn = True
            .Destination = crptToWindow
            .PrintReport
            If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
        End With
'    Else
'        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbInformation
'        Exit Sub
'    End If
'    Set RS = Nothing
End Sub

Private Sub Form_Load()
    If Check1.Value = 0 Then
        Frame1.Visible = False
    ElseIf Check1.Value = 1 Then
        Frame1.Visible = True
    End If
    If Check3.Value = 0 Then
        Frame2.Visible = False
    ElseIf Check3.Value = 1 Then
        Frame2.Visible = True
    End If
    If Check4.Value = 0 Then
        Frame3.Visible = False
    ElseIf Check4.Value = 1 Then
        Frame3.Visible = True
    End If
    
    PopulateCompany cboCompany, cboIndex
End Sub

Private Sub txtSale_KeyPress(KeyAscii As Integer)
    If txtSale.Text = "*" Then Set curStaff_ = New BWGCompanyStaff
    If KeyAscii = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If txtSale.Text = "*" Then Set curStaff_ = New BWGCompanyStaff: Exit Sub
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
