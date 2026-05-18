VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSendSampleStable 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F123"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11520
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSendBillNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "== AUTO =="
      Top             =   330
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "รายละเอียดใบส่งตัวอย่าง"
      Height          =   1965
      Left            =   30
      TabIndex        =   19
      Top             =   750
      Width           =   11445
      Begin VB.ComboBox cboSampleType 
         Height          =   315
         Left            =   8430
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1050
         Width           =   2835
      End
      Begin VB.TextBox txtSampleSetName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7740
         TabIndex        =   4
         Top             =   660
         Width           =   3525
      End
      Begin VB.TextBox txtReportNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   9450
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "== AUTO =="
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkIncTrans 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ส่ง LAB ภายนอก"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   990
         TabIndex        =   8
         Top             =   1530
         Width           =   1815
      End
      Begin VB.TextBox txtAddr 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "140 หมู่ 8 ต.ห้วยแห้ง อ.แก่งคอย จ.สระบุรี"
         Top             =   660
         Width           =   5505
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3525
      End
      Begin VB.TextBox txtLABVender 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3840
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1470
         Visible         =   0   'False
         Width           =   4485
      End
      Begin VB.TextBox txtIssuedStaff 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1050
         Width           =   2565
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "บ. เบตเตอร์เวิลด์กรีน จำกัด (มหาชน)"
         Top             =   240
         Width           =   4125
      End
      Begin WasteManagment.ctlDate dtWorkDate 
         Height          =   315
         Left            =   6555
         TabIndex        =   1
         Top             =   240
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อชุดตัวอย่าง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   6750
         TabIndex        =   31
         Top             =   720
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report No"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   8610
         TabIndex        =   30
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LAB Vender"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   2820
         TabIndex        =   28
         Top             =   1530
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ที่อยู่"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   630
         TabIndex        =   27
         Top             =   720
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "วันที่ส่งตัวอย่าง"
         Height          =   195
         Index           =   0
         Left            =   5460
         TabIndex        =   24
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ผู้ส่งตัวอย่าง"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   4860
         TabIndex        =   22
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ลูกค้า"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   1290
         TabIndex        =   21
         Top             =   270
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "บริษัทที่ส่ง"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   20
         Top             =   1140
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9540
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8340
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10530
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8340
      Width           =   945
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8340
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8340
      Width           =   945
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "รายการตัวอย่าง"
      Height          =   5430
      Left            =   30
      TabIndex        =   17
      Top             =   2730
      Width           =   11445
      Begin VB.TextBox txtRem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1020
         TabIndex        =   11
         Top             =   4950
         Width           =   10305
      End
      Begin MSComctlLib.ListView lvList 
         Height          =   4575
         Left            =   150
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   270
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8070
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "พารามิเตอร์"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "หมายเหตุ"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   4980
         Width           =   675
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8340
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบส่งตัวอย่าง"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   8370
      TabIndex        =   26
      Top             =   390
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "บาท"
      Height          =   195
      Index           =   5
      Left            =   11430
      TabIndex        =   23
      Top             =   6300
      Width           =   1395
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   90
      X2              =   7560
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ใบส่งตัวอย่างกากปรับเสถียร"
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
      TabIndex        =   18
      Top             =   90
      Width           =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   180
      X2              =   7650
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   13105
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frSendSampleStable.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmSendSampleStable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colLABParam_ As Collection
Dim curSaleStaff_ As BWGCompanyStaff
Dim mySendSample_ As BWGSendSampleHeader
Dim curCustomer_ As BWGCustomer
Dim curLABVender_ As BWGTransportVendor
Dim colSampleType_ As Collection
Dim isDetailChange_ As Boolean
Dim curSendSampleDet_ As BWGSendSampleDetail


Private Sub cboCompany_Validate(Cancel As Boolean)
    If cboCompany.ListIndex >= 0 Then
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex = 0 Then Exit Sub
            mySendSample_.companyID = CurrentUser.WorkCompanies(cboCompany.ListIndex).ID
        Else
            mySendSample_.companyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
        End If
    Else
        mySendSample_.companyID = ""
    End If
End Sub

Private Sub cboSampleType_Validate(Cancel As Boolean)
    If cboSampleType.ListIndex > 0 Then
        mySendSample_.IntSampleTypeID = colSampleType_(cboSampleType.ListIndex).ID
        mySendSample_.IntSampleTypeName = Trim(cboSampleType.Text)
    Else
        mySendSample_.IntSampleTypeID = ""
        mySendSample_.IntSampleTypeName = ""
    End If
End Sub

Private Sub chkIncTrans_Validate(Cancel As Boolean)
    If chkIncTrans.Value = Checked Then
        mySendSample_.IsTestExtByVender = True
    Else
        mySendSample_.IsTestExtByVender = False
    End If
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = mySendSample_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            txtCustomer.SetFocus
        Else
            MsgBox tmpStr, vbExclamation
        End If
    End If
End Sub
Private Sub SetEnabledScreen()
Dim eFlag As Boolean
'    eFlag = Not mySendSample_.Delete
'    cmdApproved.Enabled = eFlag
'    cmdAddItem.Enabled = eFlag
'    cmdSave.Enabled = eFlag
'    cmdCancel.Enabled = eFlag
'    Command1.Enabled = eFlag
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If mySendSample_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If mySendSample_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewSendSample
    dtWorkDate.SetFocus
End Sub

Private Sub ClearScreen()
Dim x%
    txtSendBillNo.Text = "== AUTO =="
    dtWorkDate.ValueYMD = TodayDate
    'txtCustomer.Text = ""
    'txtAddr.Text = ""
    Set curCustomer_ = Nothing
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    Set curSaleStaff_ = Nothing
    txtRem.Text = ""
    txtLABVender.Text = ""
    chkIncTrans.Value = 0
    txtReportNo.Text = "== AUTO =="
    txtSampleSetName.Text = ""
    cboSampleType.ListIndex = 0
    For x = 1 To lvList.ListItems.Count
        lvList.ListItems(x).Checked = False
    Next
End Sub

Private Sub cmdSave_Click()
Dim x%
Dim tmpDet As BWGSendSampleDetail
    If Trim(txtCustomer.Text) = "" Then
        MsgBox "กรุณาระบุชื่อลูกค้า", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    If cboCompany.ListIndex < 0 Then
        MsgBox "กรุณาระบุบริษัทที่ส่ง", vbExclamation
        cboCompany.SetFocus
        Exit Sub
    End If

    If chkIncTrans.Value = Checked Then
        If Trim(txtLABVender.Text) = "" Then
            MsgBox "กรุณาระบุLAB Vender", vbExclamation
            txtLABVender.SetFocus
            Exit Sub
        End If
    End If
    If Trim(txtSampleSetName.Text) = "" Then
        MsgBox "กรุณาระบุชื่อชุดตัวอย่าง", vbExclamation
        txtSampleSetName.SetFocus
        Exit Sub
    End If
    If cboSampleType.ListIndex = 0 Then
        MsgBox "กรุณาระบุประเภทตัวอย่างของเสีย", vbExclamation
        cboSampleType.SetFocus
        Exit Sub
    End If
    If mySendSample_.isChanged Then
        mySendSample_.Save
        txtReportNo.Text = mySendSample_.ReportNo
        txtSendBillNo.Text = mySendSample_.SendBillNo
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
Dim x%
Dim tmpKey As String
Dim tmpVal As String
Dim tmpDet As BWGSendSampleDetail
Dim tmpCol As Collection
    If mySendSample_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If mySendSample_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmSendSampleIntSearch
    With frmSendSampleIntSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Public Sub SetDataSendSample(selSend As BWGSendSampleHeader)
Dim tmpStr
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection

    Set mySendSample_ = selSend
    With mySendSample_
        txtSendBillNo.Text = .SendBillNo
        dtWorkDate.ValueYMD = .SendDate
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1), False
        'txtAddr.Text = .CustomerName
        txtIssuedStaff.Text = .IssuedStaffName
        If .IsTestExtByVender Then
            chkIncTrans.Value = 1
        Else
            chkIncTrans.Value = 0
        End If
        Set tmpCol = tmpS.TransportVenderSearch("VendorID='" & .LABVenderID & "'")
        If tmpCol.Count > 0 Then SetDataVenderLAB tmpCol(1), False
        'txtLABVender = .LABVenderName
        txtRem.Text = .RemDesc
        txtReportNo.Text = .ReportNo
        txtSampleSetName.Text = .SampleSetName
        SetTextToCombo .IntSampleTypeName, cboSampleType
        Set curSendSampleDet_ = .SendSampleDetList(1)
        ShowSendSampleDetails
    End With
    Set tmpCol = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set mySendSample_ = Nothing
Set curCustomer_ = Nothing
Set curLABVender_ = Nothing
End Sub

Private Sub lvList_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If mySendSample_.SendSampleDetList Is Nothing Then
        MsgBox "กรุณษฌาเลือกรายการของเสีย"
        Exit Sub
    End If
    If mySendSample_.SendSampleDetList.Count = 0 Then
        MsgBox "กรุณาเลือกรายการของเสีย"
        Exit Sub
    End If
    If Item.Checked Then
        If isExist(curSendSampleDet_.SampleParamReq, Item.Key) = False Then
            curSendSampleDet_.SampleParamReq.Add colLABParam_(Item.Key), Item.Key
        End If
    Else
        If isExist(curSendSampleDet_.SampleParamReq, Item.Key) Then
            curSendSampleDet_.SampleParamReq.Remove Item.Key
        End If
    End If
    curSendSampleDet_.isChanged = True
    mySendSample_.isChanged = True
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
'    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
'    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
'        Set curCustomer_ = Nothing
'        Set curCustomer_ = New BWGCustomer
'        With mySendSample_
'            .CustomerID = ""
'            .CustomerName = ""
'        End With
'    Else
'        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
'            Set curCustomer_ = Nothing
'            Set curCustomer_ = New BWGCustomer
'            With mySendSample_
'                .CustomerID = ""
'                .CustomerName = ""
'            End With
'            Load frmCustomerSearch
'            With frmCustomerSearch
'                .getCriteria (txtCustomer.Text)
'                .SetCallerForm Me
'                .Show 1
'            End With
'            If mySendSample_.CustomerID = "" Then txtCustomer.Text = ""
'        End If
'    End If
End Sub

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
Dim x%
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    With curCustomer_
        txtCustomer.Text = .CustomerName
        txtAddr.Text = .Address & " " & .TumbolName & " " & .AumphurName & " " & .ProvinceName & " " & .ZipCode
        If isAssigned Then
            mySendSample_.CustomerID = .ID
            mySendSample_.CustomerName = .CustomerName
        End If
    End With
End Sub

Private Sub chkIncTrans_Click()
    If chkIncTrans.Value = Checked Then
        Label2.Item(4).Visible = True
        txtLABVender.Visible = True
        'txtLABVender.SetFocus
    Else
        Label2.Item(4).Visible = False
        txtLABVender.Visible = False
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub ShowLABParamList()
Dim x%
Dim tmpList As ListItem
Dim tmpItem As BWGLABParameterList
Dim searchM As New BWGSearchManager
    
    Set colLABParam_ = searchM.ParameterLABSearch
    
    If colLABParam_ Is Nothing Then Exit Sub
    lvList.ListItems.Clear
    For x = 1 To colLABParam_.Count
        Set tmpItem = colLABParam_(x)
        Set tmpList = lvList.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.ParameterName)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
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
        cmdSearch.Picture = .picSearch.Picture
        cmdClear.Picture = .picNew.Picture
        cmdCancel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
'    PopulateSamplingPoint
    PopulateSampleType
    ShowLABParamList
    PopulateCompany cboCompany, cboIndex
    dtWorkDate.ValueYMD = TodayDate
    txtIssuedStaff.Text = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
    CreateNewSendSample
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Me.ActiveControl.Name <> "grSendSample" Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub CreateNewSendSample()
    Set mySendSample_ = Nothing
    Set mySendSample_ = New BWGSendSampleHeader
    Set curSendSampleDet_ = Nothing
    Set curSendSampleDet_ = New BWGSendSampleDetail
    
    With mySendSample_
        .ID = GetGUID
        .ReportNo = "int"
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                .companyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                .companyID = CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID
            End If
        End If
        .SendDate = TodayDate
        .IssuedStaffID = CurrentUser.ID
        .SendStaffID = CurrentUser.ID
        .SampleRecStaffID = CurrentUser.ID
        .IssuedStaffName = CurrentUser.StaffFName & " " & CurrentUser.StaffLName
        .IsTestExtByVender = False
        .HeaderType = "C"
        .isChanged = False
    End With
    
    With curSendSampleDet_
        .ID = GetGUID
        .SendSampleID = mySendSample_.ID
        .SamplePointName = "ปรับเสถียร"
        .isChanged = False
        mySendSample_.SendSampleDetList.Add curSendSampleDet_, "'" & .ID & "'"
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If mySendSample_ Is Nothing Then Exit Sub
    If mySendSample_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            'cmdSave_Click
            If mySendSample_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub txtLABVender_Validate(Cancel As Boolean)
    If curLABVender_ Is Nothing Then Set curLABVender_ = New BWGTransportVendor
    If Trim(txtLABVender.Text) = "" And Trim(curLABVender_.NameTH) <> Trim(txtLABVender.Text) Then
        Set curLABVender_ = Nothing
        Set curLABVender_ = New BWGTransportVendor
        With mySendSample_
            .LABVenderID = ""
            .LABVenderName = ""
        End With
    Else
        If Trim(curLABVender_.NameTH) <> Trim(txtLABVender.Text) Then
            Set curLABVender_ = Nothing
            Set curLABVender_ = New BWGTransportVendor
            With mySendSample_
                .LABVenderID = ""
                .LABVenderName = ""
            End With
            Load frmLABVenderSearch
            With frmLABVenderSearch
                .getCriteria (txtLABVender.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If mySendSample_.LABVenderID = "" Then txtLABVender.Text = ""
        End If
    End If
End Sub

Public Sub SetDataVenderLAB(selTransportVender As BWGTransportVendor, Optional isAssigned As Boolean = True)
Dim x%
    Set curLABVender_ = Nothing
    Set curLABVender_ = selTransportVender
    If curLABVender_ Is Nothing Then Exit Sub
    With curLABVender_
        txtLABVender.Text = .NameTH
        If isAssigned Then
            mySendSample_.LABVenderID = .ID
            mySendSample_.LABVenderName = .NameTH
        End If
    End With
End Sub

Private Sub txtRem_Validate(Cancel As Boolean)
    mySendSample_.RemDesc = txtRem.Text
End Sub

Private Sub ShowSendSampleDetails()
Dim x%, Y%
Dim tmpKey As String, rowCount As Long
Dim forColor As String
    For x = 1 To curSendSampleDet_.SampleParamReq.Count
        lvList.ListItems("'" & curSendSampleDet_.SampleParamReq(x).ID & "'").Checked = isExist(lvList.ListItems, "'" & curSendSampleDet_.SampleParamReq(x).ID & "'")
    Next
End Sub

Private Sub PopulateSampleType()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colSampleType_ = tmpS.InternalSampleTypeSearch
    With cboSampleType
        .Clear
        .AddItem "== กรุณาระบุประเภทตัวอย่างของเสีย =="
        For x = 1 To colSampleType_.Count
            .AddItem colSampleType_(x).TypeName
        Next
        .ListIndex = 0
    End With
End Sub

Private Sub txtSampleSetName_Validate(Cancel As Boolean)
    mySendSample_.SampleSetName = Trim(txtSampleSetName.Text)
End Sub
