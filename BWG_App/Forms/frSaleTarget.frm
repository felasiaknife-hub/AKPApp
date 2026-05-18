VERSION 5.00
Begin VB.Form frmSaleTarget 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "การตั้งเป้ายอดขาย - F116"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   810
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   120
      Width           =   2865
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   6945
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5010
      Width           =   945
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   7935
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5010
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   9915
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5010
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   8925
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5010
      Width           =   945
   End
   Begin WasteManagment.ctlGrid grTarget 
      Height          =   4335
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   10935
      _ExtentX        =   17066
      _ExtentY        =   7646
   End
   Begin VB.TextBox txtSaleDistrict 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6420
      MaxLength       =   80
      TabIndex        =   3
      Top             =   120
      Width           =   3105
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4380
      MaxLength       =   80
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VB.Line Line2 
      X1              =   -2640
      X2              =   10935
      Y1              =   4950
      Y2              =   4950
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "เขตพื้นขาย"
      Height          =   195
      Index           =   0
      Left            =   5610
      TabIndex        =   4
      Top             =   180
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ปี พ.ศ."
      Height          =   195
      Index           =   8
      Left            =   3840
      TabIndex        =   2
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "บริษัท"
      Height          =   195
      Index           =   2
      Left            =   330
      TabIndex        =   0
      Top             =   180
      Width           =   420
   End
End
Attribute VB_Name = "frmSaleTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curDistrict_ As BWGSaleDistrict
Dim colTarget_ As Collection
Dim isChanged_ As Boolean
Dim isAP As Boolean

Private Sub SetGrid(ByVal AKP As Boolean, Optional ByVal MiniView As Long = 1000)
    With grTarget.getGridObj
        .Cols = 11
        .Rows = 13
        If AKP = False Then
            .TextMatrix(0, 1) = "เดือน"
            .TextMatrix(0, 2) = "สถานะ"
            .TextMatrix(0, 3) = "รวมยอด AKP"
            .TextMatrix(0, 4) = "เป้า Non"
            .TextMatrix(0, 5) = "เป้า Hz"
            .TextMatrix(0, 6) = "เป้า NBL"
            .TextMatrix(0, 7) = "เป้า HBL"
            .TextMatrix(0, 8) = "รวม"
            .TextMatrix(0, 9) = "เป้า KPI"
            .TextMatrix(0, 10) = "Last Update"
        Else
            .TextMatrix(0, 1) = "เดือน"
            .TextMatrix(0, 2) = "สถานะ"
            .TextMatrix(0, 3) = "รวมยอด BWG"
            .TextMatrix(0, 4) = "เป้า HIC"
            .TextMatrix(0, 5) = "เป้า HBI"
            .TextMatrix(0, 6) = "เป้า BBH"
            .TextMatrix(0, 7) = "เป้า BBL"
            .TextMatrix(0, 8) = "รวม"
            .TextMatrix(0, 9) = "เป้า KPI"
            .TextMatrix(0, 10) = "Last Update"
        End If
        .ColWidth(0) = 0
        .ColWidth(1) = 675
        .ColWidth(2) = 1100
        .ColWidth(3) = 1100
        .ColWidth(4) = MiniView
        .ColWidth(5) = MiniView
        .ColWidth(6) = MiniView
        .ColWidth(7) = MiniView
        .ColWidth(8) = 1200
        .ColWidth(9) = 1100
        .ColWidth(10) = 1500
        .FixedAlignment(0) = 4
        .ColAlignment(1) = 4
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 7
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 7
        .ColAlignment(9) = 7
        .ColAlignment(10) = 4
        .TextMatrix(1, 1) = "ม.ค."
        .TextMatrix(2, 1) = "ก.พ."
        .TextMatrix(3, 1) = "มี.ค."
        .TextMatrix(4, 1) = "เม.ย."
        .TextMatrix(5, 1) = "พ.ค."
        .TextMatrix(6, 1) = "มิ.ย."
        .TextMatrix(7, 1) = "ก.ค."
        .TextMatrix(8, 1) = "ส.ค."
        .TextMatrix(9, 1) = "ก.ย."
        .TextMatrix(10, 1) = "ต.ค."
        .TextMatrix(11, 1) = "พ.ย."
        .TextMatrix(12, 1) = "ธ.ค."
    End With
    grTarget.ColInputType(2) = ComboBoxEnum
    grTarget.ColInputType(3) = ComboBoxEnum
End Sub

Private Sub cboCompany_Click()
    ShowTarget
End Sub

Private Sub cmdCancel_Click()
    If CboCompany.ListIndex < 0 Or Trim(txtYear.Text) = "" Or Trim(txtSaleDistrict.Text) = "" Then
        MsgBox "ไม่มีรายการที่ต้องการลบ", vbExclamation
        Exit Sub
    End If
    If MsgBox("ท่านต้องการลบเป้าการขายของบริษัท " & CboCompany.Text & " ในเขต " & txtSaleDistrict.Text & " ของปี พ.ศ. " & txtYear.Text & " หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        DBConnExc "Delete tbSaleTarget WHERE SaleDistrictID='" & curDistrict_.ID & "' And CompanyID='" & CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).ID & "' And ForYear=" & txtYear.Text
        ClearTarget
    End If
End Sub

Private Sub cmdClear_Click()
Dim aa
    If isChanged_ Then
        aa = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่", vbQuestion + vbYesNoCancel)
        If aa = vbYes Then
            CmdSave_Click
            If isChanged_ Then
                Exit Sub
            End If
        ElseIf aa = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CboCompany.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
Dim x%
    If CboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุบริษัทที่จะกำหนดเป้าการขาย", vbExclamation
        CboCompany.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtYear.Text) = False Then
        MsgBox "กรุณาระบุปี พ.ศ. ที่จะกำหนดเป้าการขาย", vbExclamation
        txtYear.SetFocus
        Exit Sub
    End If
    If Trim(txtSaleDistrict.Text) = "" Then
        MsgBox "กรุณาระบุเขตพื้นที่ขาย", vbExclamation
        txtSaleDistrict.SetFocus
        Exit Sub
    End If
    If isChanged_ Then
        DBConnExc "Delete tbSaleTargetDistrict WHERE SaleDistrictID='" & curDistrict_.ID & "' AND SaleStaffID = '" & curDistrict_.SaleStaffID & "' And CompanyID='" & CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).ID & "' And ForYear=" & txtYear.Text
        For x = 1 To colTarget_.Count
            colTarget_(x).Save
        Next
    End If
    isChanged_ = False
    MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grTarget" Then SendKeys "{Tab}"
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
    
    PopulateCompany CboCompany, cboIndex
    isChanged_ = False
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    SetGrid False, 1000
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If isChanged_ Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            If isChanged_ Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set curDistrict_ = Nothing
    Set colTarget_ = Nothing
    isChanged_ = False
End Sub

Private Sub grTarget_BeforeUpdateValue(NewValue As String, row As Integer, col As Integer, Cancel As Boolean)
    If col = 4 Or col = 5 Or col = 6 Or col = 7 Or col = 9 Then
        If Trim(NewValue) = "" Then Exit Sub
        If IsNumeric(NewValue) = False Then
            MsgBox "กรุณาใส่ข้อมูลเป็นตัวเลขเท่านั้น", vbExclamation
            Cancel = True
        End If
    End If
End Sub

Private Sub grTarget_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim tmpKey As String
    Dim tmpTarget As BWGSaleTarget
    With grTarget.getGridObj
        tmpKey = "'" & .TextMatrix(.row, 0) & "'"
        If isExist(colTarget_, tmpKey) Then
            If KeyCode = 46 And (.col = 4 Or .col = 5 Or .col = 6 Or .col = 7 Or .col = 9) And Trim(.TextMatrix(.row, .col)) <> "" Then
                If MsgBox("ท่านต้องการลบเป้าการขายของเดือน " & .TextMatrix(.row, 1) & " หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
                    Set tmpTarget = colTarget_(tmpKey)
                    tmpTarget.Delete
                    .RemoveItem .row
                End If
            End If
        End If
    End With
End Sub

Private Sub grTarget_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim Pont As POINTAPI
    Dim ret As BTN_STYLE

    If Button = vbRightButton Then
        hMenu = CreatePopupMenu()
        AppendMenu hMenu, MF_STRING, 1, "Full View"
        AppendMenu hMenu, MF_STRING, 2, "Mini View"
        GetCursorPos Pont
        ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD, _
                                Pont.x, Pont.Y, Me.hwnd, ByVal 0&)
        DestroyMenu hMenu

        If (CboCompany.ListIndex + cboIndex) <= 0 Then Exit Sub
        If ret = 1 Then
            If CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).CompanyCode <> "AKP" Then
                SetGrid False, 1000
            Else
                SetGrid True, 1000
            End If
        ElseIf ret = 2 Then
            If CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).CompanyCode <> "AKP" Then
                SetGrid False, 0
            Else
                SetGrid True, 0
            End If
        End If
    End If
End Sub

Private Sub grTarget_RolColChanged(row As Integer, col As Integer)
    If CboCompany.ListIndex < 0 Or Trim(txtYear.Text) = "" Or Trim(txtSaleDistrict.Text) = "" Then
        grTarget.ColEnabled(col) = False
    Else
        grTarget.ColEnabled(col) = (col = 2 Or col = 3 Or col = 4 Or col = 5 Or col = 6 Or col = 7 Or col = 8 Or col = 9)
        GetGridCBO col
    End If
End Sub

Private Sub grTarget_UpdateValue(row As Integer, col As Integer)
Dim tmpVal As String, tmpKey As String
Dim tmpTarget As BWGSaleTarget
    tmpVal = Trim(grTarget.getGridObj.TextMatrix(row, col))
    tmpKey = "'" & Trim(grTarget.getGridObj.TextMatrix(row, 0)) & "'"
    grTarget.getGridObj.TextMatrix(row, col) = Format(grTarget.getGridObj.TextMatrix(row, col), "#,##0")
    If isExist(colTarget_, tmpKey) Then
        Set tmpTarget = colTarget_(tmpKey)
    Else
        Set tmpTarget = New BWGSaleTarget
        With tmpTarget
            .ID = GetGUID
            If CboCompany.ListIndex >= 0 Then .CompanyID = CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).ID
            .CompanyName = CboCompany.Text
            .ForMonth = row
            .ForYear = txtYear.Text
            .LastUpdate = GetSystemDateTimeThai.SystemDate
            If grTarget.getColObject(2).ListIndex = 0 Then
                .InPro = 0
            ElseIf grTarget.getColObject(2).ListIndex = 1 Then
                .InPro = 1
            Else
                .InPro = 0
            End If
            .IsOverall = grTarget.getColObject(3).ListIndex
            .SaleDistrictID = curDistrict_.ID
            .SaleStaffID = curDistrict_.SaleStaffID
            .SaleDistrictName = curDistrict_.DistrictName
            .SaleStaffName = curDistrict_.SaleStaffName
            .SaleZoneName = curDistrict_.ZoneName
            grTarget.getGridObj.TextMatrix(row, 0) = .ID
        End With
        colTarget_.Add tmpTarget, "'" & tmpTarget.ID & "'"
    End If
    
    tmpTarget.InPro = IIf(grTarget.getGridObj.TextMatrix(row, 2) = "ผ่านโปร", 0, 1)
    tmpTarget.IsOverall = IIf(grTarget.getGridObj.TextMatrix(row, 3) = "ยอดเฉพาะ", 0, 1)
    If CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).ID <> "776EDA5830F744A6B3551348D851BC22" Then
        If col = 4 Then
            If IsNumeric(tmpVal) Then
                tmpTarget.NonTarget = tmpVal
            Else
                tmpTarget.NonTarget = 0
            End If
        ElseIf col = 5 Then
            If IsNumeric(tmpVal) Then
                tmpTarget.HzTarget = tmpVal
            Else
                tmpTarget.HzTarget = 0
            End If
        ElseIf col = 6 Then
            If IsNumeric(tmpVal) Then
                tmpTarget.NBLTarget = tmpVal
            Else
                tmpTarget.NBLTarget = 0
            End If
        ElseIf col = 7 Then
            If IsNumeric(tmpVal) Then
                tmpTarget.HBLTarget = tmpVal
            Else
                tmpTarget.HBLTarget = 0
            End If
        ElseIf col = 8 Then
            tmpTarget.NonTarget = IIf(IsNumeric(grTarget.getGridObj.TextMatrix(row, 8)), CDbl(grTarget.getGridObj.TextMatrix(row, 8)) / 2, 0)
            tmpTarget.HzTarget = IIf(IsNumeric(grTarget.getGridObj.TextMatrix(row, 8)), CDbl(grTarget.getGridObj.TextMatrix(row, 8)) / 2, 0)
            grTarget.getGridObj.TextMatrix(row, 4) = Format(tmpTarget.NonTarget, "#,##0")
            grTarget.getGridObj.TextMatrix(row, 5) = Format(tmpTarget.NonTarget, "#,##0")
        ElseIf col = 9 Then
            If IsNumeric(tmpVal) Then
                tmpTarget.KPITarget = tmpVal
            Else
                tmpTarget.KPITarget = 0
            End If
        End If
        grTarget.getGridObj.TextMatrix(row, 8) = Format(tmpTarget.NonTarget + tmpTarget.HzTarget + tmpTarget.NBLTarget + tmpTarget.HBLTarget, "#,##0")
    Else
        If col = 4 Then
            If IsNumeric(tmpVal) Then
                tmpTarget.HICTarget = tmpVal
            Else
                tmpTarget.HICTarget = 0
            End If
        ElseIf col = 5 Then
            If IsNumeric(tmpVal) Then
                tmpTarget.HBITarget = tmpVal
            Else
                tmpTarget.HBITarget = 0
            End If
        ElseIf col = 6 Then
            If IsNumeric(tmpVal) Then
                tmpTarget.BBHTarget = tmpVal
            Else
                tmpTarget.BBHTarget = 0
            End If
        ElseIf col = 7 Then
            If IsNumeric(tmpVal) Then
                tmpTarget.BBLTarget = tmpVal
            Else
                tmpTarget.BBLTarget = 0
            End If
        ElseIf col = 8 Then
            tmpTarget.HICTarget = IIf(IsNumeric(grTarget.getGridObj.TextMatrix(row, 8)), CDbl(grTarget.getGridObj.TextMatrix(row, 8)) / 2, 0)
            tmpTarget.HBITarget = IIf(IsNumeric(grTarget.getGridObj.TextMatrix(row, 8)), CDbl(grTarget.getGridObj.TextMatrix(row, 8)) / 2, 0)
            grTarget.getGridObj.TextMatrix(row, 4) = Format(tmpTarget.NonTarget, "#,##0")
            grTarget.getGridObj.TextMatrix(row, 5) = Format(tmpTarget.NonTarget, "#,##0")
        ElseIf col = 9 Then
            If IsNumeric(tmpVal) Then
                tmpTarget.KPITarget = tmpVal
            Else
                tmpTarget.KPITarget = 0
            End If
        End If
        grTarget.getGridObj.TextMatrix(row, 8) = Format(tmpTarget.HICTarget + tmpTarget.HBITarget + tmpTarget.BBHTarget + tmpTarget.BBLTarget, "#,##0")
    End If
    
    grTarget.getGridObj.TextMatrix(row, 10) = FormatYMD_to_DMY(tmpTarget.LastUpdate, "/", "/")
    isChanged_ = True
End Sub

Private Sub txtSaleDistrict_Validate(Cancel As Boolean)
    If curDistrict_ Is Nothing Then Set curDistrict_ = New BWGSaleDistrict
    If Trim(txtSaleDistrict.Text) = "" And Trim(curDistrict_.DistrictName) <> Trim(txtSaleDistrict.Text) Then
        Set curDistrict_ = Nothing
        Set curDistrict_ = New BWGSaleDistrict
    Else
        If Trim(curDistrict_.DistrictName) <> Trim(txtSaleDistrict.Text) Then
            Set curDistrict_ = Nothing
            Set curDistrict_ = New BWGSaleDistrict
            Load frmSaleDistrictSearch
            With frmSaleDistrictSearch
                .getCriteria (txtSaleDistrict.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curDistrict_.ID = "" Then
                txtSaleDistrict.Text = ""
            End If
        End If
    End If
    ShowTarget
End Sub

Public Sub SetDataSaleDistrict(selDistrict As BWGSaleDistrict)
    Set curDistrict_ = selDistrict
    txtSaleDistrict.Text = curDistrict_.DistrictName
End Sub

Private Sub txtYear_Validate(Cancel As Boolean)
    ShowTarget
End Sub

Private Sub ShowTarget()
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpTarget As BWGSaleTarget
Dim x%
    ClearTarget
    If CboCompany.ListIndex >= 0 Then
        If CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).ID <> "776EDA5830F744A6B3551348D851BC22" Then
            isAP = False
            SetGrid False
            If CboCompany.ListIndex >= 0 And IsNumeric(txtYear.Text) And Not curDistrict_ Is Nothing Then
                Set colTarget_ = tmpS.SaleTargetSearch("CompanyID='" & CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).ID & "' And ForYear=" & Trim(txtYear.Text) & " And SaleDistrictID='" & curDistrict_.ID & "'")
                With grTarget.getGridObj
                    For x = 1 To colTarget_.Count
                        Set tmpTarget = colTarget_(x)
                        .TextMatrix(tmpTarget.ForMonth, 0) = tmpTarget.ID
                        .TextMatrix(tmpTarget.ForMonth, 2) = IIf(tmpTarget.InPro = 0, "ผ่านโปร", "ทดลองงาน")
                        .TextMatrix(tmpTarget.ForMonth, 3) = IIf(tmpTarget.IsOverall = 0, "ยอดเฉพาะ", "ยอดรวม AKP")
                        .TextMatrix(tmpTarget.ForMonth, 4) = Format(tmpTarget.NonTarget, "#,##0")
                        .TextMatrix(tmpTarget.ForMonth, 5) = Format(tmpTarget.HzTarget, "#,##0")
                        .TextMatrix(tmpTarget.ForMonth, 6) = Format(tmpTarget.NBLTarget, "#,##0")
                        .TextMatrix(tmpTarget.ForMonth, 7) = Format(tmpTarget.HBLTarget, "#,##0")
                        .TextMatrix(tmpTarget.ForMonth, 8) = Format(tmpTarget.HzTarget + tmpTarget.NonTarget + tmpTarget.NBLTarget + tmpTarget.HBLTarget, "#,##0")
                        .TextMatrix(tmpTarget.ForMonth, 9) = Format(tmpTarget.KPITarget, "#,##0")
                        .TextMatrix(tmpTarget.ForMonth, 10) = FormatYMD_to_DMY(tmpTarget.LastUpdate, "/", "/")
                        Set tmpTarget = Nothing
                    Next
                End With
            End If
        Else
            If CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                isAP = True
                SetGrid True
                If CboCompany.ListIndex >= 0 And IsNumeric(txtYear.Text) And Not curDistrict_ Is Nothing Then
                    Set colTarget_ = tmpS.SaleTargetSearch("CompanyID='" & CurrentUser.WorkCompanies(CboCompany.ListIndex + cboIndex).ID & "' And ForYear=" & Trim(txtYear.Text) & " And SaleDistrictID='" & curDistrict_.ID & "'")
                    With grTarget.getGridObj
                        For x = 1 To colTarget_.Count
                            Set tmpTarget = colTarget_(x)
                            .TextMatrix(tmpTarget.ForMonth, 0) = tmpTarget.ID
                            .TextMatrix(tmpTarget.ForMonth, 2) = IIf(tmpTarget.InPro = 0, "ผ่านโปร", "ทดลองงาน")
                            .TextMatrix(tmpTarget.ForMonth, 3) = IIf(tmpTarget.IsOverall = 0, "ยอดเฉพาะ", "ยอดรวม BWG")
                            .TextMatrix(tmpTarget.ForMonth, 4) = Format(tmpTarget.HICTarget, "#,##0")
                            .TextMatrix(tmpTarget.ForMonth, 5) = Format(tmpTarget.HBITarget, "#,##0")
                            .TextMatrix(tmpTarget.ForMonth, 6) = Format(tmpTarget.BBHTarget, "#,##0")
                            .TextMatrix(tmpTarget.ForMonth, 7) = Format(tmpTarget.BBLTarget, "#,##0")
                            .TextMatrix(tmpTarget.ForMonth, 8) = Format(tmpTarget.HICTarget + tmpTarget.HBITarget + tmpTarget.BBHTarget + tmpTarget.BBLTarget, "#,##0")
                            .TextMatrix(tmpTarget.ForMonth, 9) = Format(tmpTarget.KPITarget, "#,##0")
                            .TextMatrix(tmpTarget.ForMonth, 10) = FormatYMD_to_DMY(tmpTarget.LastUpdate, "/", "/")
                            Set tmpTarget = Nothing
                        Next
                    End With
                End If
            End If
        End If
    End If
    Set tmpS = Nothing
End Sub

Private Sub ClearTarget()
Dim x%, Y%
    With grTarget.getGridObj
        For x = 2 To .Cols - 1
            For Y = 1 To .Rows - 1
                .TextMatrix(Y, x) = ""
            Next
        Next
    End With
End Sub

Private Sub GetGridCBO(ByVal col As Integer)
    If col = 2 Then
        With grTarget.getColObject(col)
            .Clear
            .AddItem "ผ่านโปร"
            .AddItem "ทดลองงาน"
        End With
    ElseIf col = 3 Then
        With grTarget.getColObject(col)
            .Clear
            .AddItem "ยอดเฉพาะ"
            If isAP = False Then
                .AddItem "ยอดรวม AKP"
            Else
                .AddItem "ยอดรวม BWG"
            End If
        End With
    End If
End Sub

Private Sub ClearScreen()
    If CboCompany.ListCount > 0 Then
        CboCompany.ListIndex = 0
    End If
    txtYear.Text = ""
    txtSaleDistrict.Text = ""
    Set curDistrict_ = Nothing
    Set curDistrict_ = New BWGSaleDistrict
    ClearTarget
End Sub
