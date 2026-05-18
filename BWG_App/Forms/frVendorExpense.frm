VERSION 5.00
Begin VB.Form frmVendorExpense 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "บันทึกรายจ่ายอื่นๆ ของ Vendor - F146"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2130
      Width           =   1005
   End
   Begin WasteManagment.ctlDate dtBuyDate 
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   570
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.TextBox txtMoney 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Top             =   960
      Width           =   1125
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1140
      TabIndex        =   3
      Top             =   960
      Width           =   1125
   End
   Begin VB.TextBox txtRem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1140
      TabIndex        =   5
      Top             =   1350
      Width           =   6885
   End
   Begin VB.TextBox txtDriver 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Top             =   570
      Width           =   3945
   End
   Begin VB.TextBox txtVendorCar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Top             =   180
      Width           =   6885
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2130
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2130
      Width           =   1035
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2130
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2130
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   5280
      TabIndex        =   17
      Top             =   1020
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ราคารวม"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   3390
      TabIndex        =   16
      Top             =   1020
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "จำนวน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   570
      TabIndex        =   15
      Top             =   1020
      Width           =   495
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   30
      Index           =   1
      Left            =   -30
      Top             =   1980
      Width           =   10965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "หมายเหตุ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   390
      TabIndex        =   14
      Top             =   1410
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่ซื้อ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   570
      TabIndex        =   13
      Top             =   630
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายละเอียด"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   3210
      TabIndex        =   12
      Top             =   630
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor รถ"
      Height          =   195
      Index           =   24
      Left            =   270
      TabIndex        =   11
      Top             =   240
      Width           =   780
   End
End
Attribute VB_Name = "frmVendorExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myExpense_ As BWGVendorOtherExpense
Dim curVendor_ As BWGTransportVendor
Dim callerForm_ As Form

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลรายจ่ายของ Vendor รายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myExpense_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewExpense
            txtVendorCar.SetFocus
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myExpense_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myExpense_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewExpense
    txtVendorCar.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtVendorCar.Text) = "" Then
        MsgBox "กรุณาระบุข้อมูล Vendor", vbExclamation
        txtVendorCar.SetFocus
        Exit Sub
    End If
    If Trim(txtDriver.Text) = "" Then
        MsgBox "กรุณาระบุรายการที่ซื้อ", vbExclamation
        txtDriver.SetFocus
        Exit Sub
    End If
    

    If Trim(txtQty.Text) = "" Then
        MsgBox "กรุณาระบุจำนวนที่ซื้อ", vbExclamation
        txtQty.SetFocus
        Exit Sub
    End If
    
    If Trim(txtMoney.Text) = "" Then
        MsgBox "กรุณาระบุราคาที่ซื้อ", vbExclamation
        txtMoney.SetFocus
        Exit Sub
    End If

    If IsNumeric(Trim(txtQty.Text)) = False Then
        MsgBox "กรุณาใส่ปริมาณน้ำมันที่เติมเป็นตัวเลข", vbExclamation
        txtQty.SetFocus
        Exit Sub
    End If
    If IsNumeric(Trim(txtMoney.Text)) = False Then
        MsgBox "กรุณาใส่ราคาน้ำมันต่อหน่วยเป็นตัวเลข", vbExclamation
        txtMoney.SetFocus
        Exit Sub
    End If
    
    If myExpense_.isChanged Then
        myExpense_.Save
    End If
    If myExpense_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
Dim x%
    If myExpense_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myExpense_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmVendorExpenseSearch
    With frmVendorExpenseSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub dtBuyDate_Validate(Cancel As Boolean)
    myExpense_.ExpenseDate = dtBuyDate.ValueYMD
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
        cmdSearch.Picture = .picSearch.Picture
    End With
    CreateNewExpense
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub ClearScreen()
    txtVendorCar.Text = ""
    dtBuyDate.ValueYMD = TodayDate
    txtDriver.Text = ""
    txtQty.Text = ""
    txtMoney.Text = ""
    txtRem.Text = ""
    Set curVendor_ = Nothing
    Set curVendor_ = New BWGTransportVendor
End Sub

Private Sub CreateNewExpense()
    Set myExpense_ = Nothing
    Set myExpense_ = New BWGVendorOtherExpense
    With myExpense_
        .ID = GetGUID ' getNewProductID("CC")
        If Not curVendor_ Is Nothing Then
            .VendorID = curVendor_.ID
            .VendorName = curVendor_.NameTH
        End If
        .ExpenseDate = dtBuyDate.ValueYMD
        .isChanged = False
    End With
End Sub

Public Sub SetDataVendorExpense(selRecord As BWGVendorOtherExpense)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    Set myExpense_ = Nothing
    Set myExpense_ = selRecord
    If myExpense_ Is Nothing Then Exit Sub
    With myExpense_
        dtBuyDate.ValueYMD = .ExpenseDate
        txtDriver.Text = .BillDesc
        txtQty.Text = .BuyQty
        txtMoney.Text = .NetPrice
        txtRem.Text = .RemDesc
        Set tmpCol = tmpS.TransportVenderSearch("VendorID='" & .VendorID & "'")
        If tmpCol.Count > 0 Then SetDataTransportVender tmpCol(1), False
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myExpense_ Is Nothing Then Exit Sub
    If myExpense_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myExpense_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Public Sub SetDataTransportVender(selCar As BWGTransportVendor, Optional isAssigned As Boolean = True)
    Set curVendor_ = selCar
    With curVendor_
        txtVendorCar.Text = .NameTH
        If isAssigned Then
            myExpense_.VendorID = .ID
            myExpense_.VendorName = .NameTH
        End If
    End With
End Sub

Private Sub txtDriver_Validate(Cancel As Boolean)
    myExpense_.BillDesc = txtDriver.Text
End Sub

Private Sub txtMoney_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtMoney.Text)) Then
        myExpense_.NetPrice = Trim(txtMoney.Text)
    Else
        myExpense_.NetPrice = 0
    End If
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtQty.Text)) Then
        myExpense_.BuyQty = Trim(txtQty.Text)
    Else
        myExpense_.BuyQty = 0
    End If
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myExpense_.RemDesc = txtRem.Text
End Sub

Private Sub txtVendorCar_Validate(Cancel As Boolean)
Dim selStr As String
    If curVendor_ Is Nothing Then Set curVendor_ = New BWGTransportVendor
    selStr = Trim(curVendor_.NameTH)
    If Trim(txtVendorCar.Text) = "" And Trim(selStr) <> Trim(txtVendorCar.Text) Then
        Set curVendor_ = Nothing
        Set curVendor_ = New BWGTransportVendor
        With myExpense_
            .VendorID = ""
            .VendorName = ""
        End With
    Else
        If Trim(selStr) <> Trim(txtVendorCar.Text) Then
            Set curVendor_ = Nothing
            Set curVendor_ = New BWGTransportVendor
            With myExpense_
                .VendorID = ""
                .VendorName = ""
            End With
            Load frmTransportVenderSearch
            With frmTransportVenderSearch
                .getCriteria (txtVendorCar.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myExpense_.VendorID = "" Then txtVendorCar.Text = ""
        End If
    End If
End Sub
