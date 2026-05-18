VERSION 5.00
Begin VB.Form frmComExternal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "กำหนดข้อมูลค่าคอมนอก - F012"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6060
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   10350
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6060
      Width           =   1005
   End
   Begin WasteManagment.ctlGrid grComRate 
      Height          =   5445
      Left            =   30
      TabIndex        =   2
      Top             =   540
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   9604
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   810
      TabIndex        =   0
      Top             =   120
      Width           =   4425
   End
   Begin VB.Shape Shape1 
      Height          =   5505
      Left            =   0
      Top             =   510
      Width           =   12465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ลูกค้า"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   330
      TabIndex        =   1
      Top             =   180
      Width           =   375
   End
End
Attribute VB_Name = "frmComExternal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer
Dim colExtCom_ As Collection
Dim colWaste_ As Collection
Dim isChanged_ As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim x%
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาเลือกลูกค้าที่ต้องการ", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If isChanged_ Then
        DBConnExc "Delete tbExternalCom WHERE WasteDataID in (Select WasteDataID From tbWasteDataCR WHERE CustomerID='" & curCustomer_.ID & "')"
        For x = 1 To colExtCom_.Count
            colExtCom_(x).Save
        Next
        isChanged_ = False
    End If
    MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grComRate" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If

    isChanged_ = False
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    With grComRate.getGridObj
        .Cols = 7
        .ColWidth(0) = 0
        .ColWidth(1) = 3675 + 1000
        .ColWidth(2) = 1290
        .ColWidth(3) = 1260
        .ColWidth(4) = 1170
        .ColWidth(5) = 840
        '.ColWidth(6) = 1170
        .ColWidth(6) = 2730 + 170
        .TextMatrix(0, 1) = "ชื่อกากของเสีย"
        .TextMatrix(0, 2) = "Waste No"
        .TextMatrix(0, 3) = "ค่าคอม (บาท)"
        .TextMatrix(0, 4) = "หน่วย"
        .TextMatrix(0, 5) = "ค่าคอม (%)"
        .TextMatrix(0, 6) = "ผู้รับเงิน"
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 7
        .ColAlignment(4) = 1
        .ColAlignment(5) = 7
        .ColAlignment(6) = 1
'        .FixedAlignment(7) = 4
        .FixedAlignment(1) = 4
        .FixedAlignment(2) = 4
        .FixedAlignment(3) = 4
        .FixedAlignment(4) = 4
        .FixedAlignment(5) = 4
        .FixedAlignment(6) = 4
        .Rows = 50
        grComRate.ColInputType(4) = ComboBoxEnum
        grComRate.getColObject(4).AddItem "บาท/ตัน"
        grComRate.getColObject(4).AddItem "บาท/เที่ยว"
        grComRate.getColObject(4).AddItem "บาท/ถัง"
        grComRate.getColObject(4).AddItem "บาท/ลบ.ม."
'        grComRate.ColInputType(6) = ComboBoxEnum
'        grComRate.getColObject(6).AddItem "ต่อตัน"
'        grComRate.getColObject(6).AddItem "ต่อเที่ยว"

  If CurrentUser.StaffCode = "HR-01" Then
  cmdSave.Enabled = False
  End If
    
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
  If CurrentUser.StaffCode = "HR-01" Then
        MsgBox "ท่านไม่มีสิทธิปรับปรุงข้อมูล"
  Else
    If isChanged_ Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If isChanged_ Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set curCustomer_ = Nothing
    Set colExtCom_ = Nothing
    Set colWaste_ = Nothing
End Sub

Private Sub grComRate_BeforeUpdateValue(NewValue As String, Row As Integer, Col As Integer, Cancel As Boolean)
    If (Col = 3 Or Col = 5) And IsNumeric(NewValue) = False Then
        MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลข", vbExclamation
        Cancel = True
    End If
End Sub

Private Sub grComRate_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpID As String
Dim tmpKey As String
    With grComRate.getGridObj
        If KeyCode = 46 Then
            If MsgBox("ท่านต้องการลบข้อมูลการจ่ายคอมนอกรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
                tmpID = Trim(.TextMatrix(.Row, 0))
                tmpKey = "'" & tmpID & "'"
                If isExist(colExtCom_, tmpKey) Then
                    colExtCom_.Remove tmpKey
                End If
                .TextMatrix(.Row, 3) = ""
                .TextMatrix(.Row, 4) = ""
                .TextMatrix(.Row, 5) = ""
                .TextMatrix(.Row, 6) = ""
                isChanged_ = True
            End If
        End If
    End With
End Sub

Private Sub grComRate_RolColChanged(Row As Integer, Col As Integer)
    If Trim(txtCustomer.Text) = "" Or Trim(grComRate.getGridObj.TextMatrix(Row, 0)) = "" Then
        grComRate.ColEnabled(Col) = False
    Else
        grComRate.ColEnabled(Col) = (Col = 3 Or Col = 4 Or Col = 5 Or Col = 6)
        If Col = 4 Then
            With grComRate.getColObject(Col)
                .Clear
                .AddItem "บาท/ตัน"
                .AddItem "บาท/เที่ยว"
                .AddItem "บาท/ถัง"
                .AddItem "บาท/ลบ.ม."
            End With
        End If
    End If
    
End Sub

Private Sub grComRate_UpdateValue(Row As Integer, Col As Integer)
Dim tmpExtCom As BWGExternalCom
Dim tmpID As String
Dim tmpKey As String, tmpVal As String
    With grComRate.getGridObj
        tmpID = Trim(.TextMatrix(Row, 0))
        tmpKey = "'" & tmpID & "'"
        tmpVal = Trim(.TextMatrix(Row, Col))
        If isExist(colExtCom_, tmpKey) Then
            Set tmpExtCom = colExtCom_(tmpKey)
        Else
            Set tmpExtCom = New BWGExternalCom
            tmpExtCom.ID = tmpID
            tmpExtCom.WasteName = Trim(.TextMatrix(Row, 1))
            tmpExtCom.WasteNo = Trim(.TextMatrix(Row, 2))
            If Col <> 4 Then
                tmpExtCom.RateType = "A"
                .TextMatrix(Row, 4) = "บาท/ตัน"
            End If
'            If col = 5 Then
'                tmpExtCom.RateType = "A"
'                .TextMatrix(row, 6) = "ต่อตัน"
'            End If
            colExtCom_.Add tmpExtCom, tmpKey
        End If
        Select Case Col
            Case 3
                If IsNumeric(tmpVal) Then
                    tmpExtCom.ComRateAmt = tmpVal
                Else
                    tmpExtCom.ComRateAmt = 0
                End If
            Case 4
                If Trim(tmpVal) = "บาท/ตัน" Then
                    tmpExtCom.RateType = "A"
                ElseIf Trim(tmpVal) = "บาท/เที่ยว" Then
                    tmpExtCom.RateType = "B"
                ElseIf Trim(tmpVal) = "บาท/ถัง" Then
                    tmpExtCom.RateType = "C"
                ElseIf Trim(tmpVal) = "บาท/ลบ.ม." Then
                    tmpExtCom.RateType = "D"
                End If
            Case 5
                If IsNumeric(tmpVal) Then
                    tmpExtCom.ComRatePC = tmpVal
                Else
                    tmpExtCom.ComRatePC = 0
                End If
                If tmpExtCom.RateType = "" Then
                    tmpExtCom.RateType = "A"
                    .TextMatrix(Row, 6) = "ต่อตัน"
                End If
'            Case 6
'                If Trim(tmpVal) = "ต่อตัน" Then
'                    tmpExtCom.RateTypePC = "A"
'                ElseIf Trim(tmpVal) = "ต่อเที่ยว" Then
'                    tmpExtCom.RateTypePC = "B"
'                End If
            Case 6
                tmpExtCom.ComRecName = Trim(tmpVal)
        End Select
        isChanged_ = True
        Set tmpExtCom = Nothing
    End With
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
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_.ID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer)
Dim tmpS As New BWGSearchManager
Dim x%, tmpWaste As BWGWasteDataCR
Dim tmpExtCom As BWGExternalCom
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    With curCustomer_
        txtCustomer.Text = .CustomerName
        Set colWaste_ = tmpS.WasteDataCRSearch("CustomerID='" & .ID & "'")
        Set colExtCom_ = tmpS.ExternalComSearch("WasteDataID in (Select WasteDataID from tbWasteDataCR WHERE CustomerID='" & .ID & "')")
    End With
    grComRate.ClearAllData
    With grComRate.getGridObj
        .Rows = colWaste_.Count + 50
        For x = 1 To colWaste_.Count
            Set tmpWaste = colWaste_(x)
            .TextMatrix(x, 0) = tmpWaste.ID
            .TextMatrix(x, 1) = tmpWaste.WasteName
            .TextMatrix(x, 2) = tmpWaste.WasteNo
            If isExist(colExtCom_, "'" & tmpWaste.ID & "'") Then
                Set tmpExtCom = colExtCom_("'" & tmpWaste.ID & "'")
                .TextMatrix(x, 3) = tmpExtCom.ComRateAmt
                If tmpExtCom.RateType = "A" Then
                    .TextMatrix(x, 4) = "บาท/ตัน"
                ElseIf tmpExtCom.RateType = "B" Then
                    .TextMatrix(x, 4) = "บาท/เที่ยว"
                ElseIf tmpExtCom.RateType = "C" Then
                    .TextMatrix(x, 4) = "บาท/ถัง"
                ElseIf tmpExtCom.RateType = "D" Then
                    .TextMatrix(x, 4) = "บาท/ลบ.ม."
                End If
                .TextMatrix(x, 5) = tmpExtCom.ComRatePC
'                If tmpExtCom.RateTypePC = "A" Then
'                    .TextMatrix(x, 6) = "ต่อตัน"
'                ElseIf tmpExtCom.RateTypePC = "B" Then
'                    .TextMatrix(x, 6) = "ต่อเที่ยว"
'                End If
                .TextMatrix(x, 6) = tmpExtCom.ComRecName
                Set tmpExtCom = Nothing
            End If
            Set tmpWaste = Nothing
        Next
    End With
End Sub

