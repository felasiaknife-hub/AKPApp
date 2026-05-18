VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPaymentVoucher 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F064"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboVAT 
      Height          =   315
      Left            =   9900
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1620
      Width           =   1875
   End
   Begin VB.ComboBox cboWTax 
      Height          =   315
      Left            =   6900
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1620
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "รายการเงินค้างที่จะนำมาหักในงวดนี้"
      ForeColor       =   &H80000008&
      Height          =   2355
      Index           =   2
      Left            =   60
      TabIndex        =   14
      Top             =   5040
      Width           =   11715
      Begin VB.CommandButton cmdDeSelectAll 
         Caption         =   "-"
         Height          =   225
         Left            =   3030
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "+"
         Height          =   225
         Left            =   2670
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   0
         Width           =   375
      End
      Begin MSComctlLib.ListView lvOther 
         Height          =   2025
         Left            =   60
         TabIndex        =   15
         Top             =   240
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   3572
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "วันที่"
            Object.Width           =   2249
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "รายการ"
            Object.Width           =   5371
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "ยอดเงิน"
            Object.Width           =   1826
         EndProperty
      End
   End
   Begin VB.TextBox txtVendor 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   1260
      Width           =   4275
   End
   Begin WasteManagment.ctlDate dtIssueDate 
      Height          =   375
      Left            =   3660
      TabIndex        =   1
      Top             =   900
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   661
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtRem 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      MaxLength       =   80
      TabIndex        =   10
      Top             =   1980
      Width           =   10575
   End
   Begin VB.TextBox txtRecNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4020
      TabIndex        =   7
      Top             =   1620
      Width           =   1455
   End
   Begin VB.TextBox txtInvNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   1620
      Width           =   1755
   End
   Begin VB.TextBox txtBillNo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   9900
      TabIndex        =   5
      Top             =   1260
      Width           =   1875
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   3360
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   8190
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   30
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8340
      Width           =   1005
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   6390
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8340
      Width           =   1065
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   7500
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8340
      Width           =   1005
   End
   Begin VB.CommandButton cmdPayment 
      Caption         =   "บันทึกการจ่ายเงิน"
      Height          =   795
      Left            =   60
      Picture         =   "frmPaymentVoucher.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8340
      Width           =   1725
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8340
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8340
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8340
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "รายการ"
      ForeColor       =   &H80000008&
      Height          =   2595
      Index           =   0
      Left            =   60
      TabIndex        =   27
      Top             =   2460
      Width           =   11715
      Begin WasteManagment.ctlGrid grBill 
         Height          =   2295
         Left            =   60
         TabIndex        =   11
         Top             =   210
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4048
      End
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   6900
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   900
      Width           =   4875
   End
   Begin VB.TextBox txtPaymentNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   900
      Width           =   1875
   End
   Begin WasteManagment.ctlDate dtDueDate 
      Height          =   375
      Left            =   6900
      TabIndex        =   4
      Top             =   1260
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   661
      BackColor       =   -2147483633
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   4320
      TabIndex        =   52
      Top             =   7860
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายการอื่นๆ ก่อน VAT"
      Height          =   195
      Index           =   5
      Left            =   2505
      TabIndex        =   51
      Top             =   7860
      Width           =   1590
   End
   Begin VB.Label lbTotalMenifest 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   4380
      TabIndex        =   50
      Top             =   7500
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ยอดรวม"
      Height          =   195
      Index           =   23
      Left            =   3510
      TabIndex        =   49
      Top             =   7500
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT"
      Height          =   195
      Index           =   22
      Left            =   9480
      TabIndex        =   48
      Top             =   1680
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ยอดที่ต้องสั่งจ่าย"
      Height          =   195
      Index           =   14
      Left            =   9060
      TabIndex        =   47
      Top             =   7860
      Width           =   1155
   End
   Begin VB.Label lbNetAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   10410
      TabIndex        =   46
      Top             =   7860
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAX"
      Height          =   195
      Index           =   12
      Left            =   7140
      TabIndex        =   45
      Top             =   7860
      Width           =   315
   End
   Begin VB.Label lbWTaxAmt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   7620
      TabIndex        =   44
      Top             =   7860
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   195
      Index           =   10
      Left            =   7770
      TabIndex        =   43
      Top             =   1680
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ภาษีหัก ณ. ที่จ่าย"
      Height          =   195
      Index           =   9
      Left            =   5700
      TabIndex        =   42
      Top             =   1680
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เจ้าหนี้"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   660
      TabIndex        =   41
      Top             =   1380
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ"
      Height          =   195
      Index           =   7
      Left            =   420
      TabIndex        =   40
      Top             =   2040
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "กำหนดชำระ"
      Height          =   195
      Index           =   18
      Left            =   6000
      TabIndex        =   39
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบเสร็จ"
      Height          =   195
      Index           =   17
      Left            =   3060
      TabIndex        =   38
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ใบกำกับภาษี"
      Height          =   195
      Index           =   16
      Left            =   240
      TabIndex        =   37
      Top             =   1680
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขใบวางบิล"
      Height          =   195
      Index           =   15
      Left            =   8970
      TabIndex        =   36
      Top             =   1290
      Width           =   885
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
      Left            =   4950
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลือกเครื่องพิมพ์"
      Height          =   195
      Index           =   8
      Left            =   6960
      TabIndex        =   34
      Top             =   90
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   8250
      Y2              =   8250
   End
   Begin VB.Label lbTotalPrice 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   2340
      TabIndex        =   32
      Top             =   8580
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lbVat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   7590
      TabIndex        =   31
      Top             =   7500
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VAT"
      Height          =   195
      Index           =   6
      Left            =   7140
      TabIndex        =   30
      Top             =   7500
      Width           =   375
   End
   Begin VB.Label lbDAVat 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   10410
      TabIndex        =   29
      Top             =   7500
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "อื่นๆ หลัง VAT"
      Height          =   195
      Index           =   4
      Left            =   9180
      TabIndex        =   28
      Top             =   7500
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ทำรายการชำระเงินเจ้าหนี้ทั่วไป"
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
      TabIndex        =   26
      Top             =   90
      Width           =   3030
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   150
      X2              =   6960
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "จ่ายในนาม"
      Height          =   195
      Index           =   1
      Left            =   6060
      TabIndex        =   25
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่"
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   24
      Top             =   960
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบจ่าย"
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   23
      Top             =   960
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frmPaymentVoucher.frx":046B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12030
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   3
      Left            =   10290
      Top             =   7440
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   2
      Left            =   7560
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   4
      Left            =   10290
      Top             =   7800
      Width           =   1425
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   6
      Left            =   7560
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   7
      Left            =   4260
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFF80&
      Height          =   315
      Index           =   8
      Left            =   4260
      Top             =   7800
      Width           =   1455
   End
End
Attribute VB_Name = "frmPaymentVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curVendor_ As BWGTransportVendor
Dim myBillPay_ As BWGPaymentHeader
Dim colMenifest_ As Collection
Dim colExpense_ As Collection
Dim isDetailChange_ As Boolean

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            myBillPay_.companyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).ID
        Else
            myBillPay_.companyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
        End If
        myBillPay_.CompanyName = cboCompany.Text
    Else
        myBillPay_.companyID = ""
        myBillPay_.CompanyName = ""
    End If
End Sub

Private Sub cboVAT_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = "cboVAT" Then
            If cboVAT.ListIndex = 0 Then
                myBillPay_.VATRate = 0
            Else
                myBillPay_.VATRate = SystemConfig.VATRate
            End If
            SumMoneyAmt
        End If
    End If
End Sub

Private Sub cboWTax_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = "cboWTax" Then
            myBillPay_.WTax = cboWTax.Text
            SumMoneyAmt
        End If
    End If
End Sub

Private Sub cboWTax_Validate(Cancel As Boolean)
    If IsNumeric(cboWTax.Text) Then
        myBillPay_.WTax = cboWTax.Text
    Else
        myBillPay_.WTax = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("ท่านต้องการยกเลิกการทำใบจ่ายเงินนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        myBillPay_.CancelBill
        SetCancelScreen
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myBillPay_.isChanged And myBillPay_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myBillPay_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewBillPay
    txtVendor.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub SetDataTransportVender(selVendor As BWGTransportVendor, Optional isAssigned As Boolean = True)
    Set curVendor_ = selVendor
    txtVendor.Text = selVendor.NameTH
    SetTextToCombo selVendor.WTaxRate, cboWTax
    If isAssigned Then
        myBillPay_.SupplierID = curVendor_.ID
        myBillPay_.SupplierName = curVendor_.NameTH
    End If
    ShowOtherExp
End Sub

Private Sub ShowPayDetail()
Dim x%, i%, tmpDet As BWGPaymentDetail
Dim tmpCol As Collection
    Set tmpCol = myBillPay_.PaymentDetail
    grBill.ClearAllData
    With grBill.getGridObj
        .Rows = 20 + tmpCol.Count
        Dim XX%
        For x = 1 To tmpCol.Count
            Set tmpDet = tmpCol(x)
            Dim isOther1 As Boolean, isOther2 As Boolean
            For i = 1 To lvOther.ListItems.Count
                If tmpDet.ExpenseID = Replace(lvOther.ListItems(i).Key, "'", "") Then
                    lvOther.ListItems(i).Checked = True
                    isOther1 = True
                    Exit For
                End If
            Next i
            If isOther1 = False Then
                XX = XX + 1
                .TextMatrix(XX, 0) = tmpDet.ID
                .TextMatrix(XX, 1) = tmpDet.TimeTableNo
                .TextMatrix(XX, 2) = tmpDet.BillDesc
                .TextMatrix(XX, 3) = Format(tmpDet.AmtB4VAT, "#,##0.00")
                Set tmpDet = Nothing
            End If
            isOther1 = False
        Next
    End With
    SumMoneyAmt
End Sub

Private Sub ShowOtherExp()
Dim x%, tmpItem As BWGVendorOtherExpense, tmpList As ListItem
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    Set tmpCol = tmpS.VendorExpenseSearch("VendorID='" & myBillPay_.SupplierID & "' And (BillPayHeaderID is null Or BillPayHeaderID='" & myBillPay_.ID & "')")
    lvOther.ListItems.Clear
    For x = 1 To tmpCol.Count
        Set tmpItem = tmpCol(x)
        Set tmpList = lvOther.ListItems.Add(, "'" & tmpItem.ID & "'", FormatYMD_to_DMY(tmpItem.ExpenseDate, "/", "/"))
        tmpList.SubItems(1) = Trim(tmpItem.BillDesc)
        tmpList.SubItems(2) = Format(tmpItem.NetPrice, "#,##0.00")
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    Set colExpense_ = tmpCol
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub

Private Sub cmdOpen_Click()
Dim selResult As Integer
    If myBillPay_.isChanged And myBillPay_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myBillPay_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmPaytoDeptorSearch
    With frmPaytoDeptorSearch
        .PaymentType "PaymentVoucher"
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub cmdPayment_Click()
Dim selResult As Integer
    If Trim(txtVendor.Text) = "" Then
        MsgBox "กรุณาระบุชื่อเจ้าหนี้", vbExclamation
        txtVendor.SetFocus
        Exit Sub
    End If
    If myBillPay_.isChanged And myBillPay_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            cmdSave_Click
            If myBillPay_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmCheqPayment
    With frmCheqPayment
        .SetCallerForm Me
        .TotalPayments lbNetAmt
        .Show 1
    End With
End Sub

Private Sub cmdSave_Click()
    If Trim(txtVendor.Text) = "" Then
        MsgBox "กรุณาเลือกเจ้าหนี้ที่ต้องการวางบิล", vbExclamation
        txtVendor.SetFocus
        Exit Sub
    End If
    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุบริษัทที่ออกใบวางบิล", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    If myBillPay_.PaymentDetail.Count = 0 Then
        MsgBox "กรุณาระบุรายละเอียดที่จะทำการจ่ายเงิน", vbExclamation
        Exit Sub
    End If
    If myBillPay_.isChanged And myBillPay_.isCancel = False Then
        With myBillPay_
            .AmtB4VAT = lbTotalMenifest
            .VATAmt = lbVat
            .WTaxAmt = lbWTaxAmt
            .DeductB4VATAmt = Label3
            .DeductAfterVATAmt = lbDAVat
            .TotalBillAmt = lbNetAmt
            .LastUser = Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName)
            .LastUpdate = Date & " " & Time
            .Save
        End With
        txtPaymentNo.Text = myBillPay_.PaymentNo
    End If
    If myBillPay_.isChanged = False And myBillPay_.isCancel = False Then
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
    End If
End Sub

Private Sub cmdSelectAll_Click()
Dim x%
    For x = 1 To lvOther.ListItems.Count
        If lvOther.ListItems(x).Checked = False Then
            lvOther.ListItems(x).Checked = True
            lvOther_ItemCheck lvOther.ListItems(x)
        End If
    Next
End Sub

Private Sub cmdDeSelectAll_Click()
Dim x%
    For x = 1 To lvOther.ListItems.Count
        If lvOther.ListItems(x).Checked = True Then
            lvOther.ListItems(x).Checked = False
            lvOther_ItemCheck lvOther.ListItems(x)
        End If
    Next
End Sub

Private Sub dtDueDate_Validate(Cancel As Boolean)
    myBillPay_.PaymentDueDate = dtDueDate.ValueYMD
End Sub

Private Sub dtIssueDate_Validate(Cancel As Boolean)
    myBillPay_.PaymentDate = dtIssueDate.ValueYMD
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grItem" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    PopulateCompany cboCompany, cboIndex
    With grBill.getGridObj
        .Cols = 4
        .TextMatrix(0, 1) = "เลขที่เอกสาร"
        .TextMatrix(0, 2) = "รายการ"
        .TextMatrix(0, 3) = "จำนวนเงิน"
        .ColWidth(0) = 0
        .ColWidth(1) = 1500
        .ColWidth(3) = 2000
        .ColWidth(2) = .Width - (.ColWidth(1) + .ColWidth(3)) - 300
        .Rows = 50
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .ColAlignment(1) = 4
        .ColAlignment(2) = 1
        .ColAlignment(3) = 7
    End With
    With grBill
        .ColEnabled(1) = True
        .ColEnabled(2) = True
        .ColEnabled(3) = True
    End With
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdOpen.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
        cmdPrint.Picture = .picPrint
    End With
    With cboVAT
        .AddItem "ไม่คิด VAT"
        .AddItem "คิด VAT " & SystemConfig.VATRate & "%"
        .ListIndex = 0
    End With
    With cboWTax
        .AddItem "0"
        .AddItem "1"
        .AddItem "3"
        .AddItem "5"
        .ListIndex = 0
    End With
    CreateNewBillPay
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmPaytoDeptorSearch
End Sub

Private Sub grBill_RolColChanged(Row As Integer, Col As Integer)
'    If Col <> 2 Then grBill.ColEnabled(Col) = False
End Sub

Private Sub grBill_UpdateValue(Row As Integer, Col As Integer)
Dim tmpItem As BWGPaymentDetail
Dim tmpID As String, tmpVal As String
    If Col = 3 Then
        If Trim(grBill.getGridObj.TextMatrix(Row, Col)) = "" Then
            grBill.getGridObj.TextMatrix(Row, Col) = Trim(grBill.getGridObj.TextMatrix(Row, Col))
        Else
            If IsNumeric(grBill.getGridObj.TextMatrix(Row, Col)) = False Then
                grBill.getGridObj.TextMatrix(Row, Col) = ""
            Else
                grBill.getGridObj.TextMatrix(Row, Col) = Format(grBill.getGridObj.TextMatrix(Row, Col), "#,##0.00")
            End If
        End If
    End If
    
    tmpID = Trim(grBill.getGridObj.TextMatrix(Row, 0))
    tmpVal = Trim(grBill.getGridObj.TextMatrix(Row, Col))
    If isExist(myBillPay_.PaymentDetail, "'" & tmpID & "'") Then
        Set tmpItem = myBillPay_.PaymentDetail("'" & tmpID & "'")
    Else
        Set tmpItem = New BWGPaymentDetail
        With tmpItem
            .ID = GetGUID
            .PaymentHeaderID = myBillPay_.ID
            .isCancel = False
            .RowID = getMaxRowID(myBillPay_.PaymentDetail) + 1
            myBillPay_.PaymentDetail.Add tmpItem, "'" & .ID & "'"
            grBill.getGridObj.TextMatrix(Row, 0) = Trim(.ID)
        End With
    End If
    Select Case Col
        Case 1
            tmpItem.TimeTableNo = Trim(tmpVal)
        Case 2
            tmpItem.BillDesc = Trim(tmpVal)
        Case 3
            If IsNumeric(tmpVal) Then
                tmpItem.AmtB4VAT = CDbl(tmpVal)
            Else
                tmpItem.AmtB4VAT = 0
            End If
            If Trim(tmpItem.BeforeOrAfterVAT) = "" Then
                tmpItem.BeforeOrAfterVAT = "B"
'                grBill.getGridObj.TextMatrix(Row, 3) = "ก่อน VAT"
                If IsNumeric(grBill.getGridObj.TextMatrix(Row, Col)) = False Then
                    grBill.getGridObj.TextMatrix(Row, Col) = "0.00"
                End If
                grBill.getGridObj.TextMatrix(Row, Col) = Format(grBill.getGridObj.TextMatrix(Row, Col), "#,000.00")
            End If
            
'            If grBill.getColObject(Col).ListIndex = 0 Then
'                tmpItem.BeforeOrAfterVAT = "B" 'Before
'            ElseIf grBill.getColObject(Col).ListIndex = 1 Then
'                tmpItem.BeforeOrAfterVAT = "A"  'After
'            End If
    
            SumMoneyAmt
    End Select
    myBillPay_.isChanged = True
End Sub

Private Sub lvOther_DblClick()
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection, tmpItem As BWGBillPayDetail
    If lvOther.SelectedItem Is Nothing Then Exit Sub
    Set tmpItem = colExpense_(lvOther.SelectedItem.Key)
    Set tmpCol = tmpS.VendorExpenseSearch("ExpenseID=" & lvOther.SelectedItem.Key & "")
    If tmpCol.Count > 0 Then
        Load frmVendorExpense
        With frmVendorExpense
            .SetDataVendorExpense tmpCol(1)
            .Show 1
        End With
    End If
    Set tmpItem = Nothing
    Set tmpS = Nothing
    Set tmpCol = Nothing
End Sub

Private Sub lvOther_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim x%, tmpCol As Collection
Dim tmpItem As BWGPaymentDetail, tmpID As String
Dim tmpExpense As BWGVendorOtherExpense
    Set tmpCol = myBillPay_.PaymentDetail
    tmpID = Replace(Item.Key, "'", "")
    If Item.Checked Then
        Set tmpExpense = colExpense_(Item.Key)
        Set tmpItem = New BWGPaymentDetail
        With tmpItem
            .PaymentHeaderID = myBillPay_.ID
            .ID = GetGUID
            .AmtB4VAT = tmpExpense.NetPrice
            .BillDesc = tmpExpense.BillDesc
            .isCancel = False
            .ExpenseBillDesc = tmpExpense.BillDesc
            .ExpenseBuyQty = tmpExpense.BuyQty
            .ExpenseDate = tmpExpense.ExpenseDate
            .ExpenseID = tmpExpense.ID
            .ExpenseIsPaid = tmpExpense.isPaid
            .ExpenseNetPrice = tmpExpense.NetPrice
            .ExpenseVendorID = tmpExpense.VendorID
            .ExpenseVendorName = tmpExpense.VendorName
            If cboVAT.ListIndex = 1 Then
                .VATAmt = Format(.AmtB4VAT * (SystemConfig.VATRate) / 100, "0.00")
            ElseIf cboVAT.ListIndex = 0 Then
                .VATAmt = 0
            End If
            .TotalAmtToPay = .AmtB4VAT + .VATAmt
            tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
            .RowID = getMaxRowID(tmpCol) + 1
        End With
    Else
        For x = 1 To tmpCol.Count
            If Trim(tmpID) = tmpCol(x).ExpenseID Then
                tmpCol.Remove x
                Exit For
            End If
        Next
    End If
    myBillPay_.isChanged = True
    SumMoneyAmt
End Sub

Private Sub txtBillNo_Validate(Cancel As Boolean)
    myBillPay_.BillingNo = txtBillNo.Text
End Sub

Private Sub txtInvNo_Validate(Cancel As Boolean)
    myBillPay_.InvoiceNo = txtInvNo.Text
End Sub

Private Sub txtRecNo_Validate(Cancel As Boolean)
    myBillPay_.ReceiptNo = txtRecNo.Text
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myBillPay_.BillRemark = txtRem.Text
End Sub

Private Sub txtVendor_Validate(Cancel As Boolean)
Dim selStr As String
    If curVendor_ Is Nothing Then Set curVendor_ = New BWGTransportVendor
    selStr = Trim(curVendor_.NameTH)
    If Trim(txtVendor.Text) = "" And Trim(selStr) <> Trim(txtVendor.Text) Then
        Set curVendor_ = Nothing
        Set curVendor_ = New BWGTransportVendor
        With myBillPay_
            .SupplierID = ""
            .SupplierName = ""
        End With
    Else
        If Trim(selStr) <> Trim(txtVendor.Text) Then
            Set curVendor_ = Nothing
            Set curVendor_ = New BWGTransportVendor
            With myBillPay_
                .SupplierID = ""
                .SupplierName = ""
            End With
            Load frmTransportVenderSearch
            With frmTransportVenderSearch
                .Caption = "ค้นหาเจ้าหนี้การค้าทั่วไป"
                .getCriteria (txtVendor.Text)
                .VendorType "GL"
                .SetCallerForm Me
                .Show 1
            End With
            If myBillPay_.SupplierID = "" Then txtVendor.Text = ""
        End If
    End If
End Sub

Private Sub SumMoneyAmt()
Dim tmpTotal As Double
Dim tmpTax As Double
Dim tmpTaxAmt As Double, afterVATAmt As Double
Dim x%, tripCount%
    With grBill.getGridObj
        For x = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(x, .Cols - 1)) Then
                tmpTotal = tmpTotal + CDbl(.TextMatrix(x, .Cols - 1))
                tripCount = tripCount + 1
            End If
        Next
    End With
    
    lbTotalMenifest.Caption = Format(tmpTotal, "#,##0.00")
    If IsNumeric(cboWTax.Text) Then
        tmpTax = cboWTax.Text
    Else
        tmpTax = 0
    End If
    lbWTaxAmt.Caption = Format(tmpTotal * tmpTax / 100, "#,##0.00")
    tmpTotal = 0
    
    Label3.Caption = Format(tmpTotal, "#,##0.00")
    If cboVAT.ListIndex = 1 Then
        lbVat.Caption = Format((CDbl(lbTotalMenifest.Caption) + CDbl(Label3.Caption)) * 0.07, "#,##0.00")
    ElseIf cboVAT.ListIndex = 0 Then
        lbVat.Caption = "0.00"
    End If

    tmpTotal = 0
    For x = 1 To lvOther.ListItems.Count
        If lvOther.ListItems(x).Checked And IsNumeric(lvOther.ListItems(x).SubItems(2)) Then
            tmpTotal = tmpTotal + CDbl(lvOther.ListItems(x).SubItems(2))
        End If
    Next
    
    lbDAVat.Caption = Format(afterVATAmt + tmpTotal, "#,##0.00")
    tmpTotal = lbDAVat.Caption
    
    If IsNumeric(lbWTaxAmt.Caption) Then
        tmpTaxAmt = lbWTaxAmt.Caption
    End If
    
    lbNetAmt.Caption = Format(lbTotalMenifest - lbWTaxAmt + Label3 + lbVat + lbDAVat, "#,##0.00")
    
    'lbTotalPrice.Caption = Format(CDbl(lbNetPrice.Caption) / 1.07, "#,##0.00")
End Sub

Private Sub CreateNewBillPay()
    Set myBillPay_ = Nothing
    Set myBillPay_ = New BWGPaymentHeader
    With myBillPay_
        .ID = GetGUID
        .PaymentDate = dtIssueDate.ValueYMD
        .PaymentDueDate = dtDueDate.ValueYMD
        .isChanged = False
        If cboCompany.ListIndex >= 0 Then
            If cboCompany.ListCount > 1 Then
                If cboCompany.ListIndex = 0 Then Exit Sub
                myBillPay_.companyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).ID
            Else
                myBillPay_.companyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
            End If
            myBillPay_.CompanyName = cboCompany.Text
        Else
            myBillPay_.companyID = ""
            myBillPay_.CompanyName = ""
        End If

        .isChanged = False
    End With
    SetCancelScreen
    myBillPay_.isChanged = False
End Sub

Private Sub ClearScreen()
    txtPaymentNo.Text = "== AUTO =="
    dtIssueDate.ValueYMD = TodayDate
    grBill.ClearAllData
    grBill.getGridObj.Rows = 20
    txtVendor.Text = ""
    Set curVendor_ = Nothing
    Set curVendor_ = New BWGTransportVendor
    dtDueDate.ValueYMD = TodayDate
    txtBillNo.Text = ""
    txtInvNo.Text = ""
    txtRecNo.Text = ""
    cboWTax.ListIndex = 0
    cboVAT.ListIndex = 0
    txtRem.Text = ""
    lvOther.ListItems.Clear
    lbTotalMenifest.Caption = "0.00"
    lbVat.Caption = "0.00"
    lbDAVat.Caption = "0.00"
    lbWTaxAmt.Caption = "0.00"
    lbTotalPrice.Caption = "0.00"
    lbNetAmt.Caption = "0.00"
End Sub

Private Function getMaxRowID(selCol As Collection) As Integer
Dim x%, RowID%
    For x = 1 To selCol.Count
        If RowID < selCol(x).RowID Then
            RowID = selCol(x).RowID
        End If
    Next
    getMaxRowID = RowID
End Function

Public Sub SetDataPaytoDeptor(selBillPay As BWGPaymentHeader)
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    Set myBillPay_ = Nothing
    Set myBillPay_ = selBillPay
    With myBillPay_
        txtPaymentNo.Text = .PaymentNo
        dtIssueDate.ValueYMD = .PaymentDate
        SetTextToCombo .CompanyName, cboCompany
        Set tmpCol = tmpS.TransportVenderSearch("VendorID='" & .SupplierID & "'")
        If tmpCol.Count > 0 Then
            SetDataTransportVender tmpCol(1), False
        End If
        dtDueDate.ValueYMD = .PaymentDueDate
        txtBillNo.Text = .BillingNo
        txtInvNo.Text = .InvoiceNo
        txtRecNo.Text = .ReceiptNo
        SetTextToCombo .WTax, cboWTax
        If .VATRate > 0 Then
            cboVAT.ListIndex = 1
        Else
            cboVAT.ListIndex = 0
        End If
        txtRem.Text = .BillRemark
        Set tmpCol = Nothing
        Set tmpS = Nothing
        ShowPayDetail
        .isChanged = False
    End With
    SetCancelScreen
    SumMoneyAmt
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myBillPay_ Is Nothing Then Exit Sub
    If myBillPay_.isChanged And myBillPay_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myBillPay_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub SetCancelScreen()
Dim cVal As Boolean
    cVal = myBillPay_.isCancel
    lbCancel.Visible = cVal
    cmdSave.Enabled = Not cVal
    cmdCancel.Enabled = Not cVal
End Sub

Public Function GetTripControls() As Collection
Dim x%, tmpS As New BWGSearchManager
Dim tmpCol As New Collection, tmpTrip As BWGJobDataTimeTable
Dim tmpDet As BWGPaymentDetail
Dim selCol As Collection, tmpStr As String
    Set selCol = myBillPay_.PaymentDetail
    With selCol
        For x = 1 To .Count
            Set tmpDet = .Item(x)
            Set tmpTrip = New BWGJobDataTimeTable
            If Trim(tmpDet.TimeTableNo) <> "" Then
                'tmpTrip.JobDataCarID = tmpDet.JobDataCarID
                tmpTrip.CompanyCode = tmpDet.CompanyCode
                tmpTrip.TripNo = tmpDet.TripNo
                tmpTrip.TimeTableNo = tmpDet.TimeTableNo
                tmpTrip.WorkDate = tmpDet.WorkDate
                tmpTrip.TruckSubTypeDesc = tmpDet.TruckSubTypeDesc
                tmpTrip.TruckSubTypeID = tmpDet.TruckSubTypeID
                tmpTrip.TruckTypeDesc = tmpDet.TruckTypeDesc
                tmpTrip.TruckTypeID = tmpDet.TruckTypeID
                tmpTrip.CarRegisNo = tmpDet.CarRegisNo
                tmpTrip.TrailerTruckNo = tmpDet.TrailerTruckNo
                tmpTrip.DriverName = tmpDet.DriverName
                tmpTrip.FromAumphurID = tmpDet.FromAumphurID
                tmpTrip.FromAumphurName = tmpDet.FromAumphurName
                tmpTrip.FromProvinceID = tmpDet.FromProvinceID
                tmpTrip.FromProvinceName = tmpDet.FromProvinceName
                tmpTrip.ToAumphurID = tmpDet.ToAumphurID
                tmpTrip.ToAumphurName = tmpDet.ToAumphurName
                tmpTrip.ToProvinceID = tmpDet.ToProvinceID
                tmpTrip.ToProvinceName = tmpDet.ToProvinceName
                tmpTrip.TripTranFee = tmpDet.AmtB4VAT
                tmpCol.Add tmpTrip, "'" & tmpTrip.TimeTableNo & "_" & tmpTrip.TripNo & "'"
            End If
            Set tmpTrip = Nothing
            Set tmpDet = Nothing
        Next
    End With

    Set GetTripControls = tmpCol
    Set tmpCol = Nothing
End Function

Public Sub AddTripControl(selItem As BWGJobDataTimeTable)
Dim tmpCol As Collection
Dim x%, isNew As Boolean
Dim tmpID As String, tmpCarID As String, tmpTripNo As Integer
Dim tmpDetail As New BWGPaymentDetail
    Set tmpCol = myBillPay_.PaymentDetail
    
    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpCarID = selItem.TimeTableNo
    tmpTripNo = selItem.TripNo
    For x = 1 To tmpCol.Count
        If Trim(tmpCol(x).TimeTableNo) = Trim(tmpCarID) And tmpCol(x).TripNo = tmpTripNo Then
            Set tmpDetail = tmpCol(x)
            isNew = False
            Exit For
        End If
    Next

    If isNew Then
        tmpDetail.ID = GetGUID
        tmpCol.Add tmpDetail, "'" & tmpDetail.ID & "'"
    End If
    tmpDetail.PaymentHeaderID = myBillPay_.ID
'    tmpDetail.JobDataCarID = selItem.JobDataCarID
    tmpDetail.CompanyCode = selItem.CompanyCode
    tmpDetail.TripNo = selItem.TripNo
    tmpDetail.TimeTableNo = selItem.TimeTableNo
    tmpDetail.WorkDate = selItem.WorkDate
    tmpDetail.TruckSubTypeDesc = selItem.TruckSubTypeDesc
    tmpDetail.TruckSubTypeID = selItem.TruckSubTypeID
    tmpDetail.TruckTypeDesc = selItem.TruckTypeDesc
    tmpDetail.TruckTypeID = selItem.TruckTypeID
    tmpDetail.CarRegisNo = selItem.CarRegisNo
    tmpDetail.TrailerTruckNo = selItem.TrailerTruckNo
    tmpDetail.DriverName = selItem.DriverName
    tmpDetail.FromAumphurID = selItem.FromAumphurID
    tmpDetail.FromAumphurName = selItem.FromAumphurName
    tmpDetail.FromProvinceID = selItem.FromProvinceID
    tmpDetail.FromProvinceName = selItem.FromProvinceName
    tmpDetail.ToAumphurID = selItem.ToAumphurID
    tmpDetail.ToAumphurName = selItem.ToAumphurName
    tmpDetail.ToProvinceID = selItem.ToProvinceID
    tmpDetail.ToProvinceName = selItem.ToProvinceName
    tmpDetail.RowID = tmpCol.Count
    'tmpDetail.AmtB4VAT = selItem.TripEquipFee + selItem.TripMoreWorkerFee + selItem.TripOtherFee + selItem.TripTranFee
    tmpDetail.AmtB4VAT = selItem.VDEquipFee + selItem.VDWorkerFee + selItem.VDOtherFee + selItem.VDTranFee
    tmpDetail.BillDesc = "จ่ายชำระค่าขนส่งตามใบคุมเลขที่ " & selItem.TimeTableNo & " ทะเบียน " & selItem.CarRegisNo & IIf(selItem.TrailerTruckNo = "", "", " - " & selItem.TrailerTruckNo & " (พ่วง)")
    tmpDetail.TotalAmtToPay = selItem.VDEquipFee + selItem.VDWorkerFee + selItem.VDOtherFee + selItem.VDTranFee
    isDetailChange_ = True
    myBillPay_.isChanged = True
    Set tmpDetail = Nothing
    
    If myBillPay_.PaymentDetail Is Nothing Then
        Set myBillPay_.PaymentDetail = tmpCol
        isDetailChange_ = True
        myBillPay_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub

Public Sub RemoveTripControl(selID As String)
Dim tmpCol As Collection
Dim x%, tmpKey As String
    Set tmpCol = myBillPay_.PaymentDetail
    If tmpCol Is Nothing Then Exit Sub
    For x = 1 To tmpCol.Count
        If "'" & tmpCol(x).TimeTableNo & "_" & tmpCol(x).TripNo & "'" = Trim(selID) Then
            tmpKey = "'" & tmpCol(x).ID & "'"
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myBillPay_.isChanged = True
    End If
End Sub

Public Property Get RefID() As String
    RefID = myBillPay_.ID
End Property

