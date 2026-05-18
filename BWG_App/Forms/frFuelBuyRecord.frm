VERSION 5.00
Begin VB.Form frmFuelBuyRecord 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "บันทึกการซื้อน้ำมัน - F034"
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
      Left            =   5100
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
      Left            =   5100
      TabIndex        =   2
      Top             =   570
      Width           =   2925
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
      Left            =   6300
      TabIndex        =   18
      Top             =   1020
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ลิตร"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   2340
      TabIndex        =   17
      Top             =   1020
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ราคาต่อหน่วย"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   4050
      TabIndex        =   16
      Top             =   1020
      Width           =   960
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ปริมาณ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   540
      TabIndex        =   15
      Top             =   1020
      Width           =   510
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
      Caption         =   "คนขับที่นำรถมาเติม"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   3630
      TabIndex        =   12
      Top             =   630
      Width           =   1380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รถ"
      Height          =   195
      Index           =   24
      Left            =   840
      TabIndex        =   11
      Top             =   240
      Width           =   195
   End
End
Attribute VB_Name = "frmFuelBuyRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myBuyRecord_ As BWGFuelBuyRecord
Dim curVendorCar_ As BWGVenderCar
Dim callerForm_ As Form

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลการซื้อน้ำมันรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myBuyRecord_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewBuyRecord
            txtVendorCar.SetFocus
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myBuyRecord_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myBuyRecord_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewBuyRecord
    txtVendorCar.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtVendorCar.Text) = "" Then
        MsgBox "กรุณาระบุข้อมูลรถที่ซื้อน้ำมัน", vbExclamation
        txtVendorCar.SetFocus
        Exit Sub
    End If

    If Trim(txtQty.Text) = "" Then
        MsgBox "กรุณาระบุปริมาณน้ำมันที่เติม", vbExclamation
        txtQty.SetFocus
        Exit Sub
    End If
    
    If Trim(txtMoney.Text) = "" Then
        MsgBox "กรุณาระบุราคาต่อหน่วยของน้ำมันที่เติม", vbExclamation
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
    
    If myBuyRecord_.isChanged Then
        myBuyRecord_.Save
    End If
    If myBuyRecord_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
Dim x%
    If myBuyRecord_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myBuyRecord_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmFuelBuyRecordSearch
    With frmFuelBuyRecordSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub dtBuyDate_Validate(Cancel As Boolean)
    myBuyRecord_.BuyDate = dtBuyDate.ValueYMD
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
    CreateNewBuyRecord
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
    Set curVendorCar_ = Nothing
    Set curVendorCar_ = New BWGVenderCar
End Sub

Private Sub CreateNewBuyRecord()
    Set myBuyRecord_ = Nothing
    Set myBuyRecord_ = New BWGFuelBuyRecord
    With myBuyRecord_
        .ID = GetGUID ' getNewProductID("CC")
        If Not curVendorCar_ Is Nothing Then
            .VendorCarID = curVendorCar_.ID
            .VendorName = curVendorCar_.VendorName
            .CarRegisID = curVendorCar_.CarRegisID
            .CarTypeDesc = curVendorCar_.TruckTypeDesc & ", " & curVendorCar_.SubTypeDesc
            .DriverName = curVendorCar_.ContactName
        End If
        .BuyDate = dtBuyDate.ValueYMD
        .isChanged = False
    End With
End Sub

Public Sub SetDataBuyRecord(selRecord As BWGFuelBuyRecord)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    Set myBuyRecord_ = Nothing
    Set myBuyRecord_ = selRecord
    If myBuyRecord_ Is Nothing Then Exit Sub
    With myBuyRecord_
        dtBuyDate.ValueYMD = .BuyDate
        txtDriver.Text = .DriverName
        txtQty.Text = .FuelQty
        txtMoney.Text = .UnitPrice
        txtRem.Text = .RemDesc
        Set tmpCol = tmpS.VenderCarSearch("VendorTruckID='" & .VendorCarID & "'")
        If tmpCol.Count > 0 Then SetDataCarVendor tmpCol(1), False
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myBuyRecord_ Is Nothing Then Exit Sub
    If myBuyRecord_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myBuyRecord_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Public Sub SetDataCarVendor(selCar As BWGVenderCar, Optional isAssigned As Boolean = True)
    Set curVendorCar_ = selCar
    With curVendorCar_
        txtVendorCar.Text = .TruckTypeDesc & ", " & .SubTypeDesc & " ของ " & .VendorName & " เลขทะเบียน " & .CarRegisID
        If isAssigned Then
            txtDriver.Text = .ContactName
            myBuyRecord_.CarRegisID = .CarRegisID
            myBuyRecord_.CarTypeDesc = .TruckTypeDesc & ", " & .SubTypeDesc
            myBuyRecord_.DriverName = .ContactName
            myBuyRecord_.VendorCarID = .ID
            myBuyRecord_.VendorName = .VendorName
        End If
    End With
End Sub

Private Sub txtDriver_Validate(Cancel As Boolean)
    myBuyRecord_.DriverName = txtDriver.Text
End Sub

Private Sub txtMoney_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtMoney.Text)) Then
        myBuyRecord_.UnitPrice = Trim(txtMoney.Text)
    Else
        myBuyRecord_.UnitPrice = 0
    End If
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    If IsNumeric(Trim(txtQty.Text)) Then
        myBuyRecord_.FuelQty = Trim(txtQty.Text)
    Else
        myBuyRecord_.FuelQty = 0
    End If
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    myBuyRecord_.RemDesc = txtRem.Text
End Sub

Private Sub txtVendorCar_Validate(Cancel As Boolean)
Dim selStr As String
    If curVendorCar_ Is Nothing Then Set curVendorCar_ = New BWGVenderCar
    selStr = curVendorCar_.TruckTypeDesc & ", " & curVendorCar_.SubTypeDesc & " ของ " & curVendorCar_.VendorName & " เลขทะเบียน " & curVendorCar_.CarRegisID
    If Trim(txtVendorCar.Text) = "" And Trim(selStr) <> Trim(txtVendorCar.Text) Then
        Set curVendorCar_ = Nothing
        Set curVendorCar_ = New BWGVenderCar
        With myBuyRecord_
            .CarRegisID = ""
            .CarTypeDesc = ""
            .DriverName = ""
            .VendorCarID = ""
            .VendorName = ""
        End With
    Else
        If Trim(selStr) <> Trim(txtVendorCar.Text) Then
            Set curVendorCar_ = Nothing
            Set curVendorCar_ = New BWGVenderCar
            With myBuyRecord_
                .CarRegisID = ""
                .CarTypeDesc = ""
                .DriverName = ""
                .VendorCarID = ""
                .VendorName = ""
            End With
            Load frmVendorCarSearch
            With frmVendorCarSearch
                .getCriteria (txtVendorCar.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myBuyRecord_.VendorCarID = "" Then txtVendorCar.Text = ""
        End If
    End If
End Sub
