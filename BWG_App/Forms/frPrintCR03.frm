VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmPrintCR03 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F068"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11775
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboSignName 
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   6720
      Width           =   4065
   End
   Begin VB.TextBox txtCustomer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6840
      TabIndex        =   3
      Top             =   1350
      Width           =   4785
   End
   Begin WasteManagment.ctlDate dtBillDate 
      Height          =   345
      Left            =   3660
      TabIndex        =   1
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BackColor       =   -2147483633
   End
   Begin VB.CommandButton Command2 
      Caption         =   "เลือกใบ Menifest"
      Height          =   795
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7230
      Width           =   1455
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   8190
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   5250
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7230
      Width           =   1005
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูลเก่า"
      Height          =   795
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7230
      Width           =   1065
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใบใหม่"
      Height          =   795
      Left            =   7410
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7230
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7230
      Width           =   1035
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "พิมพ์"
      Height          =   795
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7230
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7230
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ใบ Manifest ที่เลือก"
      ForeColor       =   &H80000008&
      Height          =   4845
      Index           =   0
      Left            =   30
      TabIndex        =   16
      Top             =   1770
      Width           =   11715
      Begin WasteManagment.ctlGrid grItem 
         Height          =   4575
         Left            =   60
         TabIndex        =   4
         Top             =   210
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8070
      End
   End
   Begin VB.ComboBox cboCompany 
      Height          =   315
      Left            =   6840
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   930
      Width           =   4785
   End
   Begin VB.TextBox txtBillNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   960
      Width           =   1875
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
      Caption         =   "ลูกค้า"
      Height          =   195
      Index           =   7
      Left            =   6330
      TabIndex        =   24
      Top             =   1410
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ผู้เซ็นเอกสาร"
      Height          =   195
      Index           =   6
      Left            =   150
      TabIndex        =   23
      Top             =   6780
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
      Left            =   10200
      TabIndex        =   22
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
      TabIndex        =   21
      Top             =   90
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11880
      Y1              =   7140
      Y2              =   7140
   End
   Begin VB.Label lbNetPrice 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   9420
      TabIndex        =   19
      Top             =   6750
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ตัน"
      Height          =   195
      Index           =   5
      Left            =   11430
      TabIndex        =   18
      Top             =   6750
      Width           =   225
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รวมน้ำหนักสุทธิ"
      Height          =   195
      Index           =   4
      Left            =   8100
      TabIndex        =   17
      Top             =   6750
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายงานแจ้งการกำจัดและบำบัดกากอุตสาหกรรม (CR03)"
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
      TabIndex        =   15
      Top             =   90
      Width           =   5385
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
      Caption         =   "บริษัทที่รับกำจัด"
      Height          =   195
      Index           =   1
      Left            =   5580
      TabIndex        =   14
      Top             =   990
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "วันที่"
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   13
      Top             =   1020
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่เอกสาร"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1020
      Width           =   885
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frPrintCR03.frx":0000
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
      Left            =   9300
      Top             =   6690
      Width           =   2055
   End
End
Attribute VB_Name = "frmPrintCR03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myCR03Header_ As BWGPrintCR03
Dim curSignStaff_ As BWGSystemSignPolicies
Dim curCustomer_ As BWGCustomer
Dim isDetailChange_ As Boolean
Dim colSearchResult_ As Collection
Dim SaleDistrict As String
Dim SignName As Collection

Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            myCR03Header_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).id
        Else
            myCR03Header_.CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
        End If
        myCR03Header_.CompanyName = cboCompany.Text
        SetSignPolicies myCR03Header_.CompanyID
    Else
        myCR03Header_.CompanyID = ""
        myCR03Header_.CompanyName = ""
    End If
End Sub

Private Sub cboSignName_Validate(Cancel As Boolean)
    Dim strCond As String, SignStaffID As String
    If cboSignName.ListIndex <= 0 Then Exit Sub
    Dim tmpS As New BWGSearchManager
    Set SignName = tmpS.SignNameSearch("CompanyID = '" & myCR03Header_.CompanyID & "'")
    SignStaffID = SignName(cboSignName.ListIndex).id
    If Not SignName Is Nothing Then SetDataStaff SignName(cboSignName.ListIndex)
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("ท่านต้องการยกเลิกใบแจ้งการกำจัดกากนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        myCR03Header_.CancelBill
        SetCancelScreen
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myCR03Header_.isChanged And myCR03Header_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myCR03Header_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewCR03
    dtBillDate.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
Dim selResult As Integer
    If myCR03Header_.isChanged And myCR03Header_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myCR03Header_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmPrintCR03Search
    With frmPrintCR03Search
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub cmdPrint_Click()
Dim tmpSql As String, tmpRec As New ADODB.Recordset
Dim tmpCon As ADODB.Connection, iCount%, X%, tmpCol As New Collection
Dim selResult As Integer
    If myCR03Header_.isChanged And myCR03Header_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNo)
        If selResult = vbYes Then
            CmdSave_Click
            If myCR03Header_.isChanged Then Exit Sub
        Else 'If selResult = vbCancel Then
            Exit Sub
        End If
    End If

    If cboSignName.ListIndex <= 0 Then
        MsgBox "ไม่ได้ระบุผู้เซ็นเอกสาร" & vbCrLf & "ถ้าไม่ปรากฏรายชื่อผู้มีอำนาจเซ็นเอกสาร ให้เลือกบริษัทผู้ออกเอกสารใหม่ แล้วตรวจสอบอีกครั้ง"
        Exit Sub
    End If

    With rpt1
        
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\CR03.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_CR03DetailSearch " & vbCrLf
        tmpSql = tmpSql & "WHERE CR03ID='" & myCR03Header_.id & "' "
        tmpSql = tmpSql & vbCrLf & " ORDER BY RowPos"
        .Formulas(0) = "ShowCompanyName=" & Chr(34) & myCR03Header_.CustomerName & Chr(34)
        .Formulas(1) = "ShowPrintDate=" & Chr(34) & "วันที่ " & dtBillDate.DayValue & "  เดือน " & dtBillDate.MonthNameShort & "  พ.ศ. " & dtBillDate.YearValue & Chr(34)
        .Formulas(2) = "ShowDocNo=" & Chr(34) & txtBillNo.Text & Chr(34)
        If Trim(curCustomer_.ProvinceName) = "กรุงเทพมหานคร" Then
            .Formulas(3) = "ShowAddress=" & Chr(34) & curCustomer_.Address & "  แขวง" & curCustomer_.TumbolName & "  เขต" & curCustomer_.AumphurName & "  " & curCustomer_.ProvinceName & "  " & curCustomer_.ZipCode & Chr(34)
        Else
            .Formulas(3) = "ShowAddress=" & Chr(34) & curCustomer_.Address & "  ต." & curCustomer_.TumbolName & "  อ." & curCustomer_.AumphurName & "  จ." & curCustomer_.ProvinceName & "  " & curCustomer_.ZipCode & Chr(34)
        End If
        .Formulas(4) = "ShowSignName=" & Chr(34) & curSignStaff_.SignName1 & Chr(34)
        If Not curSignStaff_ Is Nothing Then
            .Formulas(5) = "ShowSignPosition=" & Chr(34) & curSignStaff_.SignPosition1 & Chr(34)
        End If
        .Formulas(6) = "SaleDistrict=" & Chr(34) & UCase(SaleDistrict) & Chr(34)
        .Formulas(17) = "ShowUser=" & Chr(34) & CurrentUser.StaffFName & " " & CurrentUser.StaffLName & Chr(34)
        .SQLQuery = tmpSql
        .Destination = crptToWindow
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub

Private Sub CmdSave_Click()
    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุบริษัทที่ออกใบแจ้ง", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    If myCR03Header_.PrintDetails.Count = 0 Then
        MsgBox "กรุณาระบุรายละเอียดในใบแจ้ง", vbExclamation
        Exit Sub
    End If
    If myCR03Header_.isChanged And myCR03Header_.isCancel = False Then
        myCR03Header_.Save
        txtBillNo.Text = myCR03Header_.CR03No
    End If
    If myCR03Header_.isChanged = False And myCR03Header_.isCancel = False Then
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
    End If
End Sub

Private Sub Command2_Click()
    If Trim(cboCompany.Text) = "" Then
        MsgBox "กรุณาเลือกบริษัทก่อนทำการเลือกใบ Manifest", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาเลือกลูกค้าก่อนเลือกใบ Manifest", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    isDetailChange_ = False
    Load frmPrintCR03WasteAdd
    With frmPrintCR03WasteAdd
        .SetCallerForm Me, CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex), curCustomer_
        .Show 1
        If isDetailChange_ Then
            ShowPrintDetail
        End If
    End With
End Sub

Private Sub dtBilldate_Validate(Cancel As Boolean)
    myCR03Header_.IssuedDate = dtBillDate.ValueYMD
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grItem" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim X%, selX As Integer
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If

    PopulateCompany cboCompany, cboIndex
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdOpen.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
        cmdPrint.Picture = .picPrint
    End With
    With grItem.getGridObj
        .Cols = 8
        .TextMatrix(0, 1) = "วันที่ขน"
        .TextMatrix(0, 2) = "เลขที่ Manifest"
        .TextMatrix(0, 3) = "รายการของเสีย"
        .TextMatrix(0, 4) = "รหัสของเสีย"
        .TextMatrix(0, 5) = "ทะเบียนรถ"
        .TextMatrix(0, 6) = "น้ำหนัก (ตัน)"
        .TextMatrix(0, 7) = "หมายเหตุ"
        .Rows = 50
        .ColWidth(0) = 0
        .ColWidth(1) = 1020
        .ColWidth(2) = 1635
        .ColWidth(3) = 2565
        .ColWidth(4) = 1080
        .ColWidth(5) = 1260
        .ColWidth(6) = 1050
        .ColWidth(7) = 2715
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 1
        .ColAlignment(4) = 4
        .ColAlignment(5) = 1
        .ColAlignment(6) = 7
        .ColAlignment(7) = 1
    End With
    With grItem
        .ColEnabled(1) = False
        .ColEnabled(2) = False
        .ColEnabled(3) = False
        .ColEnabled(4) = False
        .ColEnabled(5) = False
        .ColEnabled(6) = False
        .ColEnabled(7) = True
    End With
    CreateNewCR03
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myCR03Header_ Is Nothing Then Exit Sub
    If myCR03Header_.isChanged And myCR03Header_.isCancel = False Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            CmdSave_Click
            If myCR03Header_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub grItem_RolColChanged(row As Integer, col As Integer)
    grItem.ColEnabled(col) = (col = 7 And Trim(grItem.getGridObj.TextMatrix(row, 0)) <> "")
End Sub

Private Sub grItem_UpdateValue(row As Integer, col As Integer)
Dim tmpDet As BWGPrintCR03Det
Dim tmpKey As String
    If col = 7 Then
        tmpKey = "'" & grItem.getGridObj.TextMatrix(row, 0) & "'"
        If isExist(myCR03Header_.PrintDetails, tmpKey) Then
            Set tmpDet = myCR03Header_.PrintDetails(tmpKey)
            tmpDet.CRRemDesc = Trim(grItem.getGridObj.Text)
            Set tmpDet = Nothing
            myCR03Header_.isChanged = True
        End If
    End If
End Sub

Private Sub ShowPrintDetail()
Dim tmpDet As BWGPrintCR03Det
Dim X%
    grItem.ClearAllData
    With grItem.getGridObj
        .Rows = myCR03Header_.PrintDetails.Count + 20
        For X = 1 To myCR03Header_.PrintDetails.Count
            Set tmpDet = myCR03Header_.PrintDetails(X)
            .TextMatrix(X, 0) = tmpDet.id
            .TextMatrix(X, 1) = FormatYMD_to_DMY(tmpDet.WorkDate)
            .TextMatrix(X, 2) = tmpDet.DocNo
            .TextMatrix(X, 3) = tmpDet.WasteName
            .TextMatrix(X, 4) = tmpDet.WasteCode
            .TextMatrix(X, 5) = tmpDet.CarRegisID
            .TextMatrix(X, 6) = tmpDet.NetWeight
            .TextMatrix(X, 7) = tmpDet.CRRemDesc
            tmpDet.RowPos = X
            Set tmpDet = Nothing
        Next
    End With
    SumTotalWeight
End Sub

Private Sub CreateNewCR03()
    Set myCR03Header_ = New BWGPrintCR03
    With myCR03Header_
        .id = GetGUID
        .IssuedDate = dtBillDate.ValueYMD
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
                .CompanyName = cboCompany.Text
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                .CompanyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).id
                .CompanyName = cboCompany.Text
            End If
        End If
        .IssuedStaffID = CurrentUser.id
        .IssuedStaffName = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        .isChanged = False
        SetCancelScreen
    End With
End Sub

Public Sub RemovePrintDetail(JobDetID As String, menifestID As String)
Dim tmpCol As Collection
Dim X%, tmpCol2 As Collection, tmpS As New BWGSearchManager
Dim tmpKey As String
    Set tmpCol = myCR03Header_.PrintDetails
    If tmpCol Is Nothing Then Exit Sub
    For X = 1 To tmpCol.Count
        If Trim(tmpCol(X).menifestID) = Trim(menifestID) And Trim(tmpCol(X).JobDetID) = Trim(JobDetID) Then
            tmpKey = "'" & tmpCol(X).id & "'"
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myCR03Header_.isChanged = True
    End If
    Set tmpS = Nothing
End Sub

Public Sub AddPrintDetail(selItem As BWGJobDataDetail)
Dim tmpCol As Collection
Dim X%, isNew As Boolean
Dim tmpID As String, treatmentDesc As String
Dim tmpDetail As New BWGPrintCR03Det, tmpS As New BWGSearchManager
    Set tmpCol = tmpS.WasteTreatmentSearch("WasteDataID='" & selItem.wastedataID & "'")
    If tmpCol.Count > 0 Then
        treatmentDesc = tmpCol(1).treatmentDesc
    Else
        treatmentDesc = ""
    End If
    Set tmpCol = myCR03Header_.PrintDetails
    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.id)
    For X = 1 To tmpCol.Count
        If tmpCol(X).menifestID = tmpID Then
            Set tmpDetail = tmpCol(X)
            isNew = False
            Exit For
        End If
    Next

    If isNew Then
        tmpDetail.id = GetGUID
        tmpDetail.CR03ID = myCR03Header_.id
        tmpCol.Add tmpDetail, "'" & tmpDetail.id & "'"
    End If
    With tmpDetail
        .CarRegisID = selItem.MenifestCarRegisID
        .DocNo = selItem.MenifestDocNo
        .JobDetID = selItem.id
        .menifestID = selItem.menifestID
        If selItem.ChargeWeightType = "A" Then
            .NetWeight = selItem.DisposerWeight
        ElseIf selItem.ChargeWeightType = "B" Then
            .NetWeight = selItem.EstWasteQty
        End If
        .RowPos = tmpCol.Count
        .treatmentDesc = treatmentDesc
        .WasteCode = selItem.WasteEUCode
        .WasteName = selItem.WasteName
        .WorkDate = selItem.MenifestWorkDate
    End With
    isDetailChange_ = True
    myCR03Header_.isChanged = True
    Set tmpDetail = Nothing
    
    If myCR03Header_.PrintDetails Is Nothing Then
        Set myCR03Header_.PrintDetails = tmpCol
        isDetailChange_ = True
        myCR03Header_.isChanged = True
    End If
    Set tmpS = Nothing
    Set tmpCol = Nothing
End Sub

Public Function GetJobDetails() As Collection
Dim X%, tmpS As New BWGSearchManager
Dim tmpCol As New Collection
Dim tmpDet As BWGPrintCR03Det, tmpJobDet As BWGJobDataDetail
Dim selCol As Collection, tmpStr As String

    Set selCol = myCR03Header_.PrintDetails
    With selCol
        For X = 1 To .Count
            Set tmpDet = .item(X)
            Set tmpJobDet = New BWGJobDataDetail
            tmpJobDet.id = tmpDet.JobDetID
            tmpJobDet.menifestID = tmpDet.menifestID
            tmpJobDet.MenifestWorkDate = tmpDet.WorkDate
            tmpJobDet.MenifestDocNo = tmpDet.DocNo
            tmpJobDet.WasteName = tmpDet.WasteName
            tmpJobDet.WasteEUCode = tmpDet.WasteCode
            tmpJobDet.MenifestCarRegisID = tmpDet.CarRegisID
            tmpJobDet.DisposerWeight = tmpDet.NetWeight
            tmpJobDet.EstWasteQty = tmpDet.NetWeight
            tmpCol.Add tmpJobDet, "'" & tmpJobDet.id & "_" & tmpJobDet.menifestID & "'"
            Set tmpDet = Nothing
            Set tmpJobDet = Nothing
        Next
    End With
    Set GetJobDetails = tmpCol
    Set tmpCol = Nothing
End Function

Private Sub ClearScreen()
    txtBillNo.Text = "== AUTO =="
    dtBillDate.ValueYMD = TodayDate
    grItem.ClearAllData
    grItem.getGridObj.Rows = 50
    lbNetPrice.Caption = "0.00"
    txtCustomer.Text = ""
    Set curCustomer_ = Nothing
    Set curCustomer_ = New BWGCustomer
    Set curSignStaff_ = Nothing
    Set curSignStaff_ = New BWGSystemSignPolicies
End Sub

Private Sub SetCancelScreen()
Dim cVal As Boolean
    cVal = myCR03Header_.isCancel
    lbCancel.Visible = cVal
    cmdSave.Enabled = Not cVal
    Command2.Enabled = Not cVal
    cmdCancel.Enabled = Not cVal
End Sub

Public Sub SetDataCR03(selBill As BWGPrintCR03)
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
Dim X%
    Set myCR03Header_ = selBill
    With myCR03Header_
        txtBillNo.Text = .CR03No
        dtBillDate.ValueYMD = .IssuedDate
        SetTextToCombo .CompanyName, cboCompany
        
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then
            SetDataCustomer tmpCol(1), False
        End If
        
        Set tmpCol = tmpS.SignNameSearch("SignStaffID='" & .SignStaffID & "'")
        If tmpCol.Count > 0 Then
            SetDataStaff tmpCol(1), False
        End If
        ShowPrintDetail
    End With
    SetCancelScreen
End Sub

Public Property Get RefID() As String
    RefID = myCR03Header_.id
End Property

Private Sub SumTotalWeight()
Dim X%, tmpAmt As Double
    With grItem.getGridObj
        For X = 1 To .Rows - 1
            If IsNumeric(.TextMatrix(X, 6)) Then
                tmpAmt = tmpAmt + CDbl(.TextMatrix(X, 6))
            End If
        Next
        lbNetPrice.Caption = Format(tmpAmt, "#,##0.000")
    End With
End Sub

Public Sub SetDataStaff(selStaff As BWGSystemSignPolicies, Optional isAssigned As Boolean = True)
    Set curSignStaff_ = selStaff
    With curSignStaff_
        SetSignPolicies myCR03Header_.CompanyID
        SetTextToCombo Trim(.SignName1), cboSignName
        If isAssigned Then
            myCR03Header_.SignStaffName = .SignName1
            myCR03Header_.SignStaffID = .id
        End If
    End With
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
        With myCR03Header_
            .CustomerID = ""
            .CustomerName = ""
        End With
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            With myCR03Header_
                .CustomerID = ""
                .CustomerName = ""
            End With
            Load frmCustomerSearch
            With frmCustomerSearch
                .getCriteria (txtCustomer.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myCR03Header_.CustomerID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim X%
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    
    With curCustomer_
        txtCustomer.Text = .CustomerNameCer
        SaleDistrict = .SaleDistrictName
        If isAssigned Then
            myCR03Header_.CustomerID = .id
            myCR03Header_.CustomerName = .CustomerNameCer
        End If
    End With
End Sub

Private Sub SetSignPolicies(ByVal CompanyID As String)
    Dim rs As New ADODB.Recordset
    Dim tSql As String, i As Integer
    tSql = "SELECT SignName1 FROM tbSystemSignPolicies WHERE CompanyID = '" & CompanyID & "' AND SignName1 IS NOT NULL  ORDER BY ConfigID"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    cboSignName.Clear
    If rs.RecordCount > 0 Then
        cboSignName.AddItem "เลือกชื่อผู้เซ็นเอกสาร"
        Do While Not rs.EOF
            cboSignName.AddItem rs!SignName1
            rs.MoveNext
        Loop
        If cboSignName.ListCount > 0 Then cboSignName.ListIndex = 0
    Else
        cboSignName.Clear
    End If
End Sub
