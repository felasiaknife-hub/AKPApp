VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdjustDetp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F156"
   ClientHeight    =   9195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9195
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSiteName 
      Height          =   315
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1710
      Width           =   10455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "เลือกรายการ"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8340
      Width           =   1245
   End
   Begin VB.TextBox txtRecPerson 
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5640
      TabIndex        =   10
      Top             =   7530
      Width           =   2475
   End
   Begin VB.TextBox txtIssueStaff 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7140
      Width           =   2475
   End
   Begin VB.TextBox txtReason 
      Appearance      =   0  'Flat
      Height          =   795
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   6300
      Width           =   4845
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8340
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8340
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8340
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใบใหม่"
      Height          =   795
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8340
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   6540
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8340
      Width           =   1065
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   5490
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8340
      Width           =   1005
   End
   Begin VB.TextBox txtCorrectAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   9960
      TabIndex        =   11
      Top             =   6690
      Width           =   1905
   End
   Begin WasteManagment.ctlGrid grItem 
      Height          =   4095
      Left            =   30
      TabIndex        =   6
      Top             =   2130
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   7223
   End
   Begin VB.ComboBox cboDocType 
      Height          =   315
      ItemData        =   "frAdjustDetp.frx":0000
      Left            =   6900
      List            =   "frAdjustDetp.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1320
      Width           =   4785
   End
   Begin VB.TextBox txtBillNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   960
      Width           =   1875
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      ItemData        =   "frAdjustDetp.frx":0086
      Left            =   6900
      List            =   "frAdjustDetp.frx":0088
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   930
      Width           =   4785
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1230
      TabIndex        =   3
      Top             =   1320
      Width           =   4275
   End
   Begin WasteManagment.ctlDate dtIssueDate 
      Height          =   345
      Left            =   3720
      TabIndex        =   1
      Top             =   930
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin MSComctlLib.ListView lvMenifest 
      Height          =   855
      Left            =   2040
      TabIndex        =   8
      Top             =   7140
      Visible         =   0   'False
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   1508
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ประเภทกลุ่มหลัก"
         Object.Width           =   4410
      EndProperty
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
      Caption         =   "สถานที่วางบิล"
      Height          =   195
      Index           =   9
      Left            =   210
      TabIndex        =   39
      Top             =   1770
      Width           =   930
   End
   Begin VB.Label lbCancel 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "ยกเลิก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   10440
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      Height          =   4155
      Left            =   0
      Top             =   2100
      Width           =   11985
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผู้รับบิล"
      Height          =   195
      Index           =   8
      Left            =   5040
      TabIndex        =   37
      Top             =   7590
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผู้ออกบิล"
      Height          =   195
      Index           =   13
      Left            =   4950
      TabIndex        =   36
      Top             =   7200
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manifest ที่เกี่ยวข้อง"
      Height          =   195
      Index           =   1
      Left            =   570
      TabIndex        =   35
      Top             =   7170
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เหตุผลในการเพิ่ม/ลด หนี้"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   6360
      Width           =   1800
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   12090
      Y1              =   8250
      Y2              =   8250
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ยอดเงินรวมทั้งสิ้น"
      Height          =   195
      Index           =   7
      Left            =   8640
      TabIndex        =   33
      Top             =   7920
      Width           =   1245
   End
   Begin VB.Label lbTotalAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   10080
      TabIndex        =   32
      Top             =   7920
      Width           =   1725
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ภาษีมูลค่าเพิ่ม"
      Height          =   195
      Index           =   6
      Left            =   8910
      TabIndex        =   31
      Top             =   7530
      Width           =   975
   End
   Begin VB.Label lbVATAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   10080
      TabIndex        =   30
      Top             =   7530
      Width           =   1725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "มูลค่าที่ถูกต้อง"
      Height          =   195
      Index           =   4
      Left            =   8940
      TabIndex        =   29
      Top             =   6750
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผลแตกต่าง"
      Height          =   195
      Index           =   5
      Left            =   9120
      TabIndex        =   28
      Top             =   7140
      Width           =   780
   End
   Begin VB.Label lbDifAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   10080
      TabIndex        =   27
      Top             =   7140
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "มูลค่าสินค้าตามใบแจ้งค่าบริการเดิมรวม"
      Height          =   195
      Index           =   19
      Left            =   7140
      TabIndex        =   26
      Top             =   6360
      Width           =   2745
   End
   Begin VB.Label lbOldAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   10080
      TabIndex        =   25
      Top             =   6360
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทเอกสาร"
      Height          =   195
      Index           =   4
      Left            =   5730
      TabIndex        =   24
      Top             =   1380
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   1020
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่"
      Height          =   195
      Index           =   2
      Left            =   3300
      TabIndex        =   22
      Top             =   1020
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ออกในนาม"
      Height          =   195
      Index           =   1
      Left            =   6000
      TabIndex        =   21
      Top             =   990
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   720
      TabIndex        =   20
      Top             =   1380
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   150
      X2              =   6960
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   7050
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เพิ่ม/ลด หนี้"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   180
      TabIndex        =   19
      Top             =   90
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frAdjustDetp.frx":008A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12030
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   1
      Left            =   9960
      Top             =   6300
      Width           =   1905
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   0
      Left            =   9960
      Top             =   7080
      Width           =   1905
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   2
      Left            =   9960
      Top             =   7470
      Width           =   1905
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   3
      Left            =   9960
      Top             =   7860
      Width           =   1905
   End
End
Attribute VB_Name = "frmAdjustDetp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer
Dim myAdjDept_ As BWGAdjDeptHeader
Dim isDetailChange_ As Boolean

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            myAdjDept_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).ID
        Else
            myAdjDept_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
        End If
        myAdjDept_.CompanyName = cboCompany.Text
    Else
        myAdjDept_.CompanyID = ""
        myAdjDept_.CompanyName = ""
    End If
End Sub

Private Sub cboDocType_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = "cboDocType" Then
            SumAllPrice False
        End If
    End If
End Sub

Private Sub cboDocType_Validate(Cancel As Boolean)
    Select Case cboDocType.ListIndex
        Case -1
            myAdjDept_.AdjustType = ""
        Case 0
            myAdjDept_.AdjustType = "A"
        Case 1
            myAdjDept_.AdjustType = "B"
        Case 2
            myAdjDept_.AdjustType = "C"
        Case 3
            myAdjDept_.AdjustType = "D"
    End Select
End Sub

Private Sub cboSiteName_Validate(Cancel As Boolean)
    If cboSiteName.ListIndex = -1 Then
        myAdjDept_.BillAddr = ""
        myAdjDept_.BillAddrID = ""
    Else
        myAdjDept_.BillAddr = cboSiteName.Text
        myAdjDept_.BillAddrID = curCustomer_.BillAddresses(cboSiteName.ListIndex + 1).ID
    End If
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("ท่านต้องการยกเลิกใบเพิ่ม/ลด หนี้รายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        myAdjDept_.CancelBill
        SetCancelScreen
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myAdjDept_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myAdjDept_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewAdjDept
    dtIssueDate.SetFocus
    SetCancelScreen
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
Dim selResult As Integer
    If myAdjDept_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myAdjDept_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmAdjustDeptSearch
    With frmAdjustDeptSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub cmdPrint_Click()
Dim tmpSql As String, tmpCompany As BWGCompanyConfig
Dim selResult As Integer, tmpAddr As BWGCustomerBillAddr
    If Trim(TxtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        TxtCustomer.SetFocus
        Exit Sub
    End If
    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุชื่อบริษัทที่ออกบิล", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If

    If myAdjDept_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            cmdSave_Click
            If myAdjDept_.isChanged Then Exit Sub
        Else
            Exit Sub
        End If
    End If
    Set tmpCompany = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex)
    With Rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\AdjustDept.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_AdjDeptDetSearch " & vbCrLf
        tmpSql = tmpSql & "WHERE AdjustID='" & myAdjDept_.ID & "' And AdjBillDesc<>'' "
        tmpSql = tmpSql & vbCrLf & " ORDER BY RowPos"
        
        .Formulas(0) = "ShowBillType=" & Chr(34) & cboDocType.Text & Chr(34)
        .Formulas(1) = "ShowCompanyTH=" & Chr(34) & tmpCompany.CompanyNameTH & Chr(34)
        .Formulas(2) = "ShowCompanyEN=" & Chr(34) & tmpCompany.CompanyNameEN & Chr(34)
        .Formulas(3) = "AddrTH1=" & Chr(34) & tmpCompany.AddrTH1 & Chr(34)
        .Formulas(4) = "AddrTH2=" & Chr(34) & tmpCompany.AddrTH2 & " " & tmpCompany.AddrTH3 & Chr(34)
        .Formulas(5) = "AddrEN1=" & Chr(34) & tmpCompany.AddrEN1 & Chr(34)
        .Formulas(6) = "AddrEN2=" & Chr(34) & tmpCompany.AddrEN2 & " " & tmpCompany.AddrEN3 & Chr(34)
        .Formulas(7) = "TelNo=" & Chr(34) & "TEL : " & tmpCompany.PhoneNo & " FAX : " & tmpCompany.FaxNo & Chr(34)
        
        .Formulas(8) = "BillNo=" & Chr(34) & txtBillNo.Text & Chr(34)
        .Formulas(9) = "TaxID=" & Chr(34) & tmpCompany.TaxID & Chr(34)
        .Formulas(10) = "IssuedDate=" & Chr(34) & dtIssueDate.DayValue & " " & dtIssueDate.MonthName & " " & dtIssueDate.YearValue & Chr(34)
        
        
        
        If cboSiteName.ListIndex = -1 Then
            .Formulas(11) = "CustomerName=" & Chr(34) & curCustomer_.CustomerName & " เลขประจำตัวผู้เสียภาษี " & tmpAddr.TaxNum & Chr(34)
            .Formulas(12) = "CustAddr1=" & Chr(34) & "" & Chr(34)
            .Formulas(13) = "CustAddr2=" & Chr(34) & "" & Chr(34)
        Else
            Set tmpAddr = curCustomer_.BillAddresses(cboSiteName.ListIndex + 1)
            .Formulas(11) = "CustomerName=" & Chr(34) & tmpAddr.BillName & " เลขประจำตัวผู้เสียภาษี " & tmpAddr.TaxNum & Chr(34)
            If InStr(1, tmpAddr.ProvinceName, "กรุงเทพ") > 0 Then
                .Formulas(12) = "CustAddr1=" & Chr(34) & tmpAddr.Address & " แขวง " & tmpAddr.TumbolName & Chr(34)
                .Formulas(13) = "CustAddr2=" & Chr(34) & " เขต " & tmpAddr.AumphurName & " " & tmpAddr.ProvinceName & " " & tmpAddr.ZipCode & Chr(34)
            Else
                .Formulas(12) = "CustAddr1=" & Chr(34) & tmpAddr.Address & " ต. " & tmpAddr.TumbolName & Chr(34)
                .Formulas(13) = "CustAddr2=" & Chr(34) & " อ. " & tmpAddr.AumphurName & " จ. " & tmpAddr.ProvinceName & " " & tmpAddr.ZipCode & Chr(34)
            End If
        End If
'                    .Formulas(2) = "CustName=" & Chr(34) & Trim(tmpSite.BillName) & Chr(34)
'                    If InStr(1, tmpSite.ProvinceName, "กรุงเทพ") > 0 Then
'                        .Formulas(3) = "CustAddr=" & Chr(34) & tmpSite.Address & " แขวง " & tmpSite.TumbolName & " เขต " & tmpSite.AumphurName & " " & tmpSite.ProvinceName & " " & tmpSite.ZipCode & Chr(34)
'                    Else
'                        .Formulas(3) = "CustAddr=" & Chr(34) & tmpSite.Address & " ต. " & tmpSite.TumbolName & " อ. " & tmpSite.AumphurName & " จ. " & tmpSite.ProvinceName & " " & tmpSite.ZipCode & Chr(34)
'                    End If
        
        If cboDocType.ListIndex = 0 Or cboDocType.ListIndex = 1 Then
            .Formulas(14) = "RemTitle=" & Chr(34) & "เหตุผลในการลดหนี้" & Chr(34)
            .Formulas(21) = "RecDocType=" & Chr(34) & "ผู้รับใบลดหนี้" & Chr(34)
            .Formulas(22) = "IssueDocType=" & Chr(34) & "ผู้ออกใบลดหนี้" & Chr(34)
        ElseIf cboDocType.ListIndex = 2 Or cboDocType.ListIndex = 3 Then
            .Formulas(14) = "RemTitle=" & Chr(34) & "เหตุผลในการเพิ่มหนี้" & Chr(34)
            .Formulas(21) = "RecDocType=" & Chr(34) & "ผู้รับใบเพิ่มหนี้" & Chr(34)
            .Formulas(22) = "IssueDocType=" & Chr(34) & "ผู้ออกใบเพิ่มหนี้" & Chr(34)
        End If
        .Formulas(15) = "RemDesc=" & Chr(34) & txtReason.Text & Chr(34)
        .Formulas(16) = "TotalAmt=" & CDbl(lbOldAmt.Caption)
        If IsNumeric(txtCorrectAmt.Text) Then
            .Formulas(17) = "CorrectAmt=" & CDbl(txtCorrectAmt.Text)
        Else
            .Formulas(17) = "CorrectAmt=" & "0"
        End If
        .Formulas(18) = "DiffAmt=" & CDbl(lbDifAmt.Caption)
        .Formulas(19) = "VATAmt=" & CDbl(lbVATAmt.Caption)
        .Formulas(20) = "SumAmtWord=" & Chr(34) & AmountToString(CDbl(lbTotalAmt)) & Chr(34)
        
        .Formulas(23) = "ShowD=" & Chr(34) & Format(dtIssueDate.DayValue, "00") & Chr(34)
        .Formulas(24) = "ShowM=" & Chr(34) & Format(dtIssueDate.MonthValue, "00") & Chr(34)
        .Formulas(25) = "ShowY=" & Chr(34) & Left(dtIssueDate.ValueYMD, 4) & Chr(34)
        
        .SQLQuery = tmpSql
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With

End Sub

Private Sub cmdSave_Click()
    If Trim(TxtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        TxtCustomer.SetFocus
        Exit Sub
    End If
    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุชื่อบริษัทที่ออกบิล", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    If myAdjDept_.AdjDetails.Count = 0 Then
        MsgBox "กรุณาระบุรายการที่ต้องการเพิ่ม/ลด หนี้", vbExclamation
        Exit Sub
    End If
    With myAdjDept_
        If .isChanged Then
            If IsNumeric(lbOldAmt.Caption) Then
                .OldBillAmt = lbOldAmt.Caption
            Else
                .OldBillAmt = 0
            End If
            .Save
            txtBillNo.Text = .AdjustBillNo
        End If
        If .isChanged = False Then
            MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
        End If
    End With
End Sub

Private Sub Command1_Click()
    If Trim(TxtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        TxtCustomer.SetFocus
        Exit Sub
    End If
    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุบริษัทที่ออกบิล", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    isDetailChange_ = False

    isDetailChange_ = False
    Load frmAdjDeptInvAdd
    With frmAdjDeptInvAdd
        .SetCallerForm Me, curCustomer_, myAdjDept_.ID, CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
        .Show 1
        If isDetailChange_ Then
            ShowAdjustDetail
        End If
    End With
End Sub

Private Sub dtIssueDate_Validate(Cancel As Boolean)
    myAdjDept_.IssuedDate = dtIssueDate.ValueYMD
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    With grItem.getGridObj
        .Cols = 7
        .Rows = 50
        .TextMatrix(0, 1) = "Invoice"
        .TextMatrix(0, 2) = "รายการ"
        .TextMatrix(0, 3) = "ปริมาณ"
        .TextMatrix(0, 4) = "หน่วย"
        .TextMatrix(0, 5) = "หน่วยละ"
        .TextMatrix(0, 6) = "จำนวนเงิน"
        .ColWidth(0) = 0
        .ColWidth(1) = 1245
        .ColWidth(2) = 5490
        .ColWidth(3) = 1080
        .ColWidth(4) = 1110
        .ColWidth(5) = 1230
        .ColWidth(6) = 1515
    End With
    Label1(6).Caption = "ภาษีมูลค่าเพิ่ม " & SystemConfig.VATRate
    txtIssueStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdOpen.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
        cmdPrint.Picture = .picPrint
    End With

    grItem.ColInputType(4) = ComboBoxEnum
    With grItem.getColObject(4)
        .AddItem "เที่ยว"
        .AddItem "ตัน"
        .AddItem "เดือน"
        .AddItem "ถัง"
        .AddItem "ลิตร"
        .AddItem "หน่วย"
    End With
    PopulateCompany cboCompany, cboIndex
    cboDocType.ListIndex = 0
    CreateNewAdjDept
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myAdjDept_ Is Nothing Then Exit Sub
    If myAdjDept_.isChanged And cmdSave.Enabled Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myAdjDept_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmAdjustDeptSearch
End Sub

Private Sub grItem_BeforeUpdateValue(NewValue As String, Row As Integer, Col As Integer, Cancel As Boolean)
    If (Col = 3 Or Col = 5) And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลข", vbExclamation
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub grItem_Click(Row As Integer, Col As Integer)
'    MsgBox grItem.getGridObj.ColWidth(col)
End Sub


Private Sub CreateNewAdjDept()
    Set myAdjDept_ = New BWGAdjDeptHeader
    With myAdjDept_
        .ID = GetGUID
        .IssuedDate = dtIssueDate.ValueYMD
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
                .CompanyName = cboCompany.Text
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
                .CompanyName = cboCompany.Text
            End If
        End If
        .IssuedStaffID = CurrentUser.ID
        .VATRate = SystemConfig.VATRate
        Select Case cboDocType.ListIndex
            Case -1
                .AdjustType = ""
            Case 0
                .AdjustType = "A"
            Case 1
                .AdjustType = "B"
            Case 2
                .AdjustType = "C"
            Case 3
                .AdjustType = "D"
        End Select
        If Not curCustomer_ Is Nothing Then
            .CustomerID = curCustomer_.ID
            .CustomerName = curCustomer_.CustomerName
        End If
        .isChanged = False
    End With
End Sub

Private Sub grItem_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpKey As String, x%, tmpCol As Collection
Dim isModify As Boolean
    With grItem.getGridObj
        If KeyCode = 45 Then 'Insert
                .AddItem "", .Row
                isModify = True
        ElseIf KeyCode = 46 Then 'Del
            If Trim(.TextMatrix(.Row, 0)) = "" Then
                .RemoveItem .Row
                .Rows = .Rows + 1
                isModify = True
            Else
                If MsgBox("ท่านต้องการลบรายการนี้ออกจากใบเพิ่ม/ลด หนี้ใบนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
                    tmpKey = "'" & .TextMatrix(.Row, 0) & "'"
                    If isExist(myAdjDept_.AdjDetails, tmpKey) Then myAdjDept_.AdjDetails.Remove tmpKey
                    .RemoveItem .Row
                    .Rows = .Rows + 1
                    myAdjDept_.isChanged = True
                    isModify = True
                End If
            End If
        End If
        If (KeyCode = 45 Or KeyCode = 46) And isModify Then
            Set tmpCol = myAdjDept_.AdjDetails
            For x = 1 To .Rows - 1
                If isExist(tmpCol, "'" & .TextMatrix(x, 0) & "'") Then
                    tmpCol("'" & .TextMatrix(x, 0) & "'").RowPos = x
                End If
            Next
        End If
    End With
End Sub

Private Sub grItem_RolColChanged(Row As Integer, Col As Integer)
    With grItem.getGridObj
        If Col = 1 Or Col = 6 Or ((Col = 3 Or Col = 4 Or Col = 5) And Trim(.TextMatrix(Row, 1)) <> "") Then
            grItem.ColEnabled(Col) = False
        ElseIf Col = 2 And Trim(.TextMatrix(Row - 1, Col)) <> "" Then
            grItem.ColEnabled(Col) = True
        Else
            If Trim(.TextMatrix(Row, 2)) <> "" Then
                grItem.ColEnabled(Col) = True
            Else
                grItem.ColEnabled(Col) = False
            End If
        End If
    End With
End Sub

Private Sub grItem_UpdateValue(Row As Integer, Col As Integer)
Dim tmpKey As String, tmpVal As String
Dim tmpDet As BWGAdjDeptDetail, tmpCol As Collection
    With grItem.getGridObj
        If Trim(.TextMatrix(Row, 0)) = "" Then
            .TextMatrix(Row, 0) = GetGUID
        End If
        tmpKey = "'" & .TextMatrix(Row, 0) & "'"
        tmpVal = Trim(.TextMatrix(Row, Col))
    End With
    Set tmpCol = myAdjDept_.AdjDetails
    If isExist(tmpCol, tmpKey) Then
        Set tmpDet = tmpCol(tmpKey)
    Else
        Set tmpDet = New BWGAdjDeptDetail
        tmpDet.ID = Trim(Replace(tmpKey, "'", ""))
        tmpDet.AdjustID = myAdjDept_.ID
        tmpDet.RowPos = Row
        tmpCol.Add tmpDet, tmpKey
    End If
    Select Case Col
        Case 2
            tmpDet.AdjBillDesc = tmpVal
        Case 3
            If IsNumeric(tmpVal) Then
                tmpDet.AdjBillQty = tmpVal
            Else
                tmpDet.AdjBillQty = 0
            End If
            grItem.getGridObj.TextMatrix(Row, 6) = Format(tmpDet.AdjBillQty * tmpDet.AdjUnitPrice, "#,##0.00")
        Case 4
            tmpDet.AdjUnitName = tmpVal
        Case 5
            If IsNumeric(tmpVal) Then
                tmpDet.AdjUnitPrice = tmpVal
            Else
                tmpDet.AdjUnitPrice = 0
            End If
            grItem.getGridObj.TextMatrix(Row, 6) = Format(tmpDet.AdjBillQty * tmpDet.AdjUnitPrice, "#,##0.00")
    End Select
    If tmpDet.isChanged Then myAdjDept_.isChanged = True
    Set tmpDet = Nothing
    Set tmpCol = Nothing
    SumAllPrice True
End Sub

Private Sub lbDifAmt_Change()
    With lbDifAmt
        If IsNumeric(lbDifAmt.Caption) Then
            If CDbl(lbDifAmt.Caption) < 0 Then
                .ForeColor = vbRed
                .FontBold = True
            Else
                .ForeColor = vbWhite
                .FontBold = False
            End If
        Else
            .ForeColor = vbWhite
            .FontBold = False
        End If
    End With
End Sub

Private Sub lbTotalAmt_Change()
    With lbTotalAmt
        If IsNumeric(.Caption) Then
            If CDbl(.Caption) < 0 Then
                .ForeColor = vbRed
                .FontBold = True
            Else
                .ForeColor = vbWhite
                .FontBold = False
            End If
        Else
            .ForeColor = vbWhite
            .FontBold = False
        End If
    End With
End Sub

Private Sub lbVATAmt_Change()
    With lbVATAmt
        If IsNumeric(.Caption) Then
            If CDbl(.Caption) < 0 Then
                .ForeColor = vbRed
                .FontBold = True
            Else
                .ForeColor = vbWhite
                .FontBold = False
            End If
        Else
            .ForeColor = vbWhite
            .FontBold = False
        End If
    End With
End Sub

Private Sub txtCorrectAmt_Change()
    If Me.Visible Then
        If Me.ActiveControl.Name = "txtCorrectAmt" Then
            SumAllPrice False
        End If
    End If
End Sub

Private Sub txtCorrectAmt_Validate(Cancel As Boolean)
    If IsNumeric(txtCorrectAmt.Text) Then
        myAdjDept_.CorrectBillAmt = txtCorrectAmt.Text
    Else
        myAdjDept_.CorrectBillAmt = 0
    End If
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
Dim selStr As String
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    selStr = Trim(curCustomer_.CustomerName)
    If Trim(TxtCustomer.Text) = "" And Trim(selStr) <> Trim(TxtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
        With myAdjDept_
            .CustomerID = ""
            .CustomerName = ""
        End With
    Else
        If Trim(selStr) <> Trim(TxtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            With myAdjDept_
                .CustomerID = ""
                .CustomerName = ""
            End With
            Load frmCustomerSearch
            With frmCustomerSearch
                .CStatus = "N"
                .getCriteria (TxtCustomer.Text)
                .GetCriteriaOptonal ""
                .SetCallerForm Me
                .Show 1
            End With
            If myAdjDept_.CustomerID = "" Then TxtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCust As BWGCustomer, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set curCustomer_ = selCust
    cboSiteName.Clear
    With curCustomer_
        For x = 1 To .BillAddresses.Count
            cboSiteName.AddItem .BillAddresses(x).Address & " " & .BillAddresses(x).TumbolName & " " & .BillAddresses(x).AumphurName & " " & .BillAddresses(x).ProvinceName & " " & .BillAddresses(x).ZipCode
        Next
        
        TxtCustomer.Text = .CustomerName
        If isAssigned Then
            myAdjDept_.CustomerID = .ID
            myAdjDept_.CustomerName = .CustomerName
            If cboSiteName.ListCount = 1 Then
                cboSiteName.ListIndex = 0
                myAdjDept_.BillAddrID = curCustomer_.BillAddresses(1).ID
                myAdjDept_.BillAddr = cboSiteName.Text
            End If
        End If
    End With
End Sub

'Public Sub AddBillItem(selItem As BWGAdjDeptDetail)
'Dim tmpCol As Collection
'Dim x%, isNew As Boolean
'Dim tmpID As String
'    Set tmpCol = myAdjDept_.AdjDetails
'    If tmpCol Is Nothing Then Set tmpCol = New Collection
'    isNew = Not isExist(tmpCol, "'" & selItem.ID & "'")
'
'    If isNew Then
'        selItem.AdjustID = myAdjDept_.ID
'        selItem.RowPos = tmpCol.Count + 1
'        tmpCol.Add selItem, "'" & selItem.ID & "'"
'        isDetailChange_ = True
'        myAdjDept_.isChanged = True
'    End If
'    If myAdjDept_.AdjDetails Is Nothing Then
'        Set myAdjDept_.AdjDetails = tmpCol
'        isDetailChange_ = True
'        myAdjDept_.isChanged = True
'    End If
'    Set tmpCol = Nothing
'End Sub

Public Sub AddBillItem(selItem As BWGBillingHeader)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String
Dim tmpDetail As New BWGAdjDeptDetail
    Set tmpCol = myAdjDept_.AdjDetails
    
    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.ID)
    For x = 1 To tmpCol.Count
        'If tmpCol(x).selitem.ID = tmpID Then
        If tmpCol(x).BillingHeaderID = tmpID Then
            isNew = False
            Exit For
        End If
    Next

    If isNew Then
        tmpDetail.ID = GetGUID
        tmpDetail.BillingHeaderID = selItem.ID
        tmpDetail.AdjBillDesc = "ค่าบริการบำบัดและกำจัดกากอุตสาหกรรม"
        tmpDetail.AdjBillQty = 1
        tmpDetail.AdjUnitName = "Invoice"
        tmpDetail.AdjUnitPrice = selItem.AmtB4VAT
        tmpDetail.AdjustID = myAdjDept_.ID
        tmpDetail.BillingNo = selItem.BillingNo
        tmpDetail.RowPos = tmpCol.Count + 1
        tmpCol.Add tmpDetail, "'" & tmpDetail.ID & "'"
        isDetailChange_ = True
        myAdjDept_.isChanged = True
        Set tmpDetail = Nothing
    End If
    If myAdjDept_.AdjDetails Is Nothing Then
        Set myAdjDept_.AdjDetails = tmpCol
        isDetailChange_ = True
        myAdjDept_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub

Public Sub RemoveBillData(BillHeaderID As String)
Dim tmpCol As Collection
Dim x%
Dim tmpKey As String
    Set tmpCol = myAdjDept_.AdjDetails
    If tmpCol Is Nothing Then Exit Sub
    
    For x = 1 To tmpCol.Count
        If Trim(tmpCol(x).BillingHeaderID) = Trim(BillHeaderID) Then
            tmpKey = "'" & tmpCol(x).ID & "'"
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myAdjDept_.isChanged = True
    End If
End Sub

Private Sub ShowAdjustDetail()
Dim x%, tmpCol As Collection
Dim tmpAdjDet As BWGAdjDeptDetail
    Set tmpCol = myAdjDept_.AdjDetails
    grItem.ClearAllData
    lvMenifest.ListItems.Clear
    With grItem.getGridObj
        .Rows = tmpCol.Count + 20
        For x = 1 To tmpCol.Count
            Set tmpAdjDet = tmpCol(x)
            .TextMatrix(x, 0) = tmpAdjDet.ID
            .TextMatrix(x, 1) = tmpAdjDet.BillingNo
            .TextMatrix(x, 2) = tmpAdjDet.AdjBillDesc
            If tmpAdjDet.AdjBillQty > 0 Then .TextMatrix(x, 3) = tmpAdjDet.AdjBillQty
            
            .TextMatrix(x, 4) = tmpAdjDet.AdjUnitName
            
            If tmpAdjDet.AdjUnitPrice > 0 Then .TextMatrix(x, 5) = tmpAdjDet.AdjUnitPrice
            If tmpAdjDet.AdjBillQty * tmpAdjDet.AdjUnitPrice > 0 Then .TextMatrix(x, 6) = Format(tmpAdjDet.AdjBillQty * tmpAdjDet.AdjUnitPrice, "#,##0.00")
            If Trim(tmpAdjDet.MenifestDocNo) <> "" And lvMenifest.FindItem(tmpAdjDet.MenifestDocNo) Is Nothing Then
                lvMenifest.ListItems.Add , "'" & tmpAdjDet.ManifestID & "'", tmpAdjDet.MenifestDocNo
            End If
            tmpAdjDet.RowPos = x
        Next
    End With
    SumAllPrice True
End Sub

Public Function getBillItems() As Collection
Dim x%
Dim tmpCol As New Collection
Dim tmpDet As BWGAdjDeptDetail
Dim tmpBill As BWGBillingHeader
Dim selCol As Collection, tmpStr As String, tmpArr
    Set selCol = myAdjDept_.AdjDetails
    With selCol
        For x = 1 To .Count
            Set tmpDet = .item(x)
            If Trim(tmpDet.BillingHeaderID) <> "" Then
                Set tmpBill = New BWGBillingHeader
                tmpBill.ID = tmpDet.BillingHeaderID
                tmpBill.BillingNo = tmpDet.BillingNo
                tmpBill.AmtB4VAT = tmpDet.AdjUnitPrice
                tmpCol.Add tmpBill, "'" & tmpDet.BillingHeaderID & "'"
                Set tmpBill = Nothing
            End If
            Set tmpDet = Nothing
        Next
    End With

    Set getBillItems = tmpCol
    Set tmpCol = Nothing
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grItem" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub SumAllPrice(reCalInGrid As Boolean)
Dim x%, amt1 As Double, amt2 As Double
    If reCalInGrid Then
        With grItem.getGridObj
            For x = 1 To .Rows - 1
                If IsNumeric(Trim(.TextMatrix(x, 6))) Then
                    amt1 = amt1 + CDbl(Trim(.TextMatrix(x, 6)))
                End If
            Next
            lbOldAmt.Caption = Format(amt1, "#,##0.00")
        End With
    End If
    If IsNumeric(lbOldAmt.Caption) Then
        amt1 = CDbl(lbOldAmt.Caption)
    Else
        amt1 = 0
    End If
    If IsNumeric(txtCorrectAmt.Text) Then
        amt2 = CDbl(txtCorrectAmt.Text)
    Else
        amt2 = 0
    End If
    If cboDocType.ListIndex = 0 Or cboDocType.ListIndex = 1 Then
        lbDifAmt.Caption = Format(amt1 - amt2, "#,##0.00")
    ElseIf cboDocType.ListIndex = 2 Or cboDocType.ListIndex = 3 Then
        lbDifAmt.Caption = Format(amt2 - amt1, "#,##0.00")
    End If
    If IsNumeric(lbDifAmt.Caption) Then
        amt1 = CDbl(lbDifAmt.Caption)
    Else
        amt1 = 0
    End If
    lbVATAmt.Caption = Format(amt1 * myAdjDept_.VATRate / 100, "#,##0.00")
    lbTotalAmt.Caption = Format(CDbl(lbVATAmt.Caption) + CDbl(lbDifAmt.Caption), "#,##0.00")
End Sub

Private Sub txtReason_Validate(Cancel As Boolean)
    myAdjDept_.AdjReason = Trim(txtReason.Text)
End Sub

Private Sub txtRecPerson_Validate(Cancel As Boolean)
    myAdjDept_.RecBillPerson = Trim(txtRecPerson.Text)
End Sub

Private Sub ClearScreen()
    txtBillNo.Text = "== AUTO =="
    dtIssueDate.ValueYMD = TodayDate
    cboCompany.ListIndex = 0
    TxtCustomer.Text = ""
    Set curCustomer_ = Nothing
    Set curCustomer_ = New BWGCustomer
    cboDocType.ListIndex = 0
    grItem.ClearAllData
    grItem.getGridObj.Rows = 50
    txtReason.Text = ""
    lvMenifest.ListItems.Clear
    txtRecPerson.Text = ""
    lbOldAmt.Caption = "0.00"
    txtCorrectAmt.Text = ""
    lbDifAmt.Caption = "0.00"
    lbVATAmt.Caption = "0.00"
    lbTotalAmt.Caption = "0.00"
    Me.cboSiteName.ListIndex = -1
    Label1(6).Caption = "ภาษีมูลค่าเพิ่ม " & SystemConfig.VATRate
End Sub

Private Sub SetCancelScreen()
Dim cVal As Boolean
    cVal = myAdjDept_.isCancel
    lbCancel.Visible = cVal
    cmdSave.Enabled = Not cVal
    cmdCancel.Enabled = Not cVal
    Command1.Enabled = Not cVal
End Sub

Public Sub SetDataAdjustDept(selAdjDept As BWGAdjDeptHeader)
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    Set myAdjDept_ = Nothing
    Set myAdjDept_ = selAdjDept
    With myAdjDept_
        txtBillNo.Text = .AdjustBillNo
        dtIssueDate.ValueYMD = .IssuedDate
        SetTextToCombo .CompanyName, cboCompany
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then
            SetDataCustomer tmpCol(1), False
        End If
        SetTextToCombo .BillAddr, cboSiteName
        If .AdjustType = "A" Then
            cboDocType.ListIndex = 0
        ElseIf .AdjustType = "B" Then
            cboDocType.ListIndex = 1
        ElseIf .AdjustType = "C" Then
            cboDocType.ListIndex = 2
        ElseIf .AdjustType = "D" Then
            cboDocType.ListIndex = 3
        End If
        txtReason.Text = .AdjReason
        txtIssueStaff.Text = .IssuedStaffName
        txtRecPerson.Text = .RecBillPerson
        txtCorrectAmt.Text = .CorrectBillAmt
        ShowAdjustDetail
    End With
    SetCancelScreen
End Sub

