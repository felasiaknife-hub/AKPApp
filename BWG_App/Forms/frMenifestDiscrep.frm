VERSION 5.00
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Begin VB.Form frmMenifestDiscrep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Discrepancy Notification - F060"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   5775
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   7950
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Width           =   1005
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   5850
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2280
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ยกเลิก"
      Height          =   795
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2280
      Width           =   1035
   End
   Begin VB.ComboBox cboWasteType 
      Height          =   315
      ItemData        =   "frMenifestDiscrep.frx":0000
      Left            =   1620
      List            =   "frMenifestDiscrep.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   540
      Width           =   1005
   End
   Begin VB.TextBox txtDiscrepQty 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4530
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   510
      Width           =   1545
   End
   Begin VB.OptionButton optAction 
      Caption         =   "ส่งคืน"
      Height          =   225
      Index           =   0
      Left            =   1650
      TabIndex        =   4
      Top             =   930
      Width           =   795
   End
   Begin VB.OptionButton optAction 
      Caption         =   "จัดประเภทใหม่"
      Height          =   225
      Index           =   1
      Left            =   2610
      TabIndex        =   5
      Top             =   930
      Width           =   1455
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4530
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   870
      Width           =   1545
   End
   Begin VB.TextBox txtRemDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7890
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   870
      Width           =   2025
   End
   Begin VB.TextBox txtReturnNo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7890
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1230
      Width           =   2025
   End
   Begin VB.TextBox txtTsdf 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1650
      TabIndex        =   11
      Top             =   1620
      Width           =   4395
   End
   Begin VB.ComboBox cboDiscrepUnit 
      Height          =   315
      Left            =   6150
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   510
      Width           =   1245
   End
   Begin VB.OptionButton optAction 
      Caption         =   "รับกำจัด"
      Height          =   225
      Index           =   2
      Left            =   6210
      TabIndex        =   7
      Top             =   930
      Width           =   915
   End
   Begin DateCtl.Date dtReturnDate 
      Height          =   315
      Left            =   1650
      TabIndex        =   9
      Top             =   1230
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายละเอียดของเสีย"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   24
      Top             =   210
      Width           =   1335
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   1
      Left            =   0
      Top             =   2100
      Width           =   10275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทของเสีย"
      Height          =   195
      Index           =   6
      Left            =   480
      TabIndex        =   23
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ปริมาณ"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   51
      Left            =   3900
      TabIndex        =   22
      Top             =   570
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "การดำเนินงาน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   52
      Left            =   510
      TabIndex        =   21
      Top             =   930
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัส"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   53
      Left            =   4170
      TabIndex        =   20
      Top             =   930
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เหตุผล"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   54
      Left            =   7290
      TabIndex        =   19
      Top             =   930
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "วันที่ส่งคืน"
      Height          =   195
      Index           =   4
      Left            =   840
      TabIndex        =   18
      Top             =   1290
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "หมายเลขใบกำกับการขนส่งของเสียอัตรายที่ส่งกลับ"
      Height          =   195
      Index           =   5
      Left            =   4380
      TabIndex        =   17
      Top             =   1290
      Width           =   3435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ลงชื่อผู้ส่งคืน"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   55
      Left            =   690
      TabIndex        =   16
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "frmMenifestDiscrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myDiscrep_ As BWGDiscrepWaste
Dim curWasteData_ As BWGWasteDataCR
Dim curTSDF_ As BWGTSDF
Dim curMenifest_ As BWGMenifest
Dim callerForm_ As Form

Private Sub cboDiscrepUnit_Validate(Cancel As Boolean)
    If cboDiscrepUnit.ListIndex > -1 Then
        myDiscrep_.QtyUnitID = colUnit_(cboDiscrepUnit.ListIndex + 1).ID
        myDiscrep_.QtyUnitName = colUnit_(cboDiscrepUnit.ListIndex + 1).UnitName
    Else
        myDiscrep_.QtyUnitID = ""
        myDiscrep_.QtyUnitName = ""
    End If
End Sub

Private Sub cboWasteType_Validate(Cancel As Boolean)
    myDiscrep_.WasteType = cboWasteType.Text
End Sub

Private Sub cmdCancel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลรายการของเสียปนเปื้อนของลูกค้ารายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        tmpStr = myDiscrep_.Delete
        If Trim(tmpStr) = "" Then
            ClearScreen
            CreateNewDiscrep
            txtDesc.SetFocus
        Else
            MsgBox tmpStr
        End If
    End If
End Sub

Private Sub cmdClear_Click()
Dim selResult As Integer
    If myDiscrep_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myDiscrep_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewDiscrep
    txtDesc.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Trim(txtDesc.Text) = "" Then
        MsgBox "กรุณาระบุรายละเอียดของเสีย", vbExclamation
        txtDesc.SetFocus
        Exit Sub
    End If

    If Trim(cboWasteType.Text) = "" Then
        MsgBox "กรุณาระบุประเภทของเสีย", vbExclamation
        cboWasteType.SetFocus
        Exit Sub
    End If

    If myDiscrep_.isChanged Then
        myDiscrep_.Save
    End If
    If myDiscrep_.isChanged = False Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
End Sub

Private Sub dtReturnDate_Validate(Cancel As Boolean)
    myDiscrep_.ReturnDate = dtReturnDate.ValueYMD
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
    End With
    CreateNewDiscrep
    PopulateUnit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub ClearScreen()
    txtDesc.Text = ""
    cboWasteType.ListIndex = -1
    txtDiscrepQty.Text = ""
    cboDiscrepUnit.ListIndex = -1
    optAction(0).Value = False
    optAction(1).Value = False
    optAction(2).Value = False
    txtCode.Text = ""
    txtRemDesc.Text = ""
    dtReturnDate.ValueYMD = TodayDate
    txtReturnNo.Text = ""
    txtTsdf.Text = ""
    Set curWasteData_ = Nothing
    Set curTSDF_ = Nothing
    Set curWasteData_ = New BWGWasteDataCR
    Set curTSDF_ = New BWGTSDF
End Sub

Private Sub CreateNewDiscrep()
    Set myDiscrep_ = Nothing
    Set myDiscrep_ = New BWGDiscrepWaste
    With myDiscrep_
        .ID = GetGUID ' getNewProductID("CC")
        If Not curMenifest_ Is Nothing Then
            .menifestID = curMenifest_.ID
            .OldMenifestDocNo = curMenifest_.docNo
            .OldMenifestNo = curMenifest_.MenifestNo
        End If
        .ReturnDate = dtReturnDate.ValueYMD
        .isChanged = False
    End With
End Sub

Public Sub SetDataMenifestDiscrep(selDiscrep As BWGDiscrepWaste)
Dim tmpCol As Collection
Dim tmpS As New BWGSearchManager
    Set myDiscrep_ = Nothing
    Set myDiscrep_ = selDiscrep
    If myDiscrep_ Is Nothing Then Exit Sub
    With myDiscrep_
        txtDesc.Text = .DiscrepDesc
        SetTextToCombo .WasteType, cboWasteType
        txtDiscrepQty.Text = .DiscrepQty
        SetTextToCombo .QtyUnitName, cboDiscrepUnit
        If .ActionTaken = "A" Then
            optAction(0).Value = True
        ElseIf .ActionTaken = "B" Then
            optAction(1).Value = True
        ElseIf .ActionTaken = "C" Then
            optAction(2).Value = True
        End If
        txtRemDesc.Text = .ReasonAction
        dtReturnDate.ValueYMD = .ReturnDate
        txtReturnNo.Text = .ReturnMenifestID
        
        'Set tmpCol = tmpS.WasteDataCRSearch("WasteDataID='" & .NewWasteDataID & "'")
        'If tmpCol.Count > 0 Then SetDataWasteData tmpCol(1), False
        
        Set tmpCol = tmpS.TSDFSearch("TSDFID='" & .TSDFID & "'")
        If tmpCol.Count > 0 Then SetDataTSDF tmpCol(1), False
        txtCode.Text = .NewWasteDataID
        
        Set tmpCol = Nothing
        Set tmpS = Nothing
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selResult As Integer
    If myDiscrep_ Is Nothing Then Exit Sub
    If myDiscrep_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myDiscrep_.isChanged Then Cancel = 1
        ElseIf selResult = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Public Sub SetCallerForm(selFrm As Form, selMenifest As BWGMenifest)
    Set callerForm_ = selFrm
    Set curMenifest_ = selMenifest
    If Not myDiscrep_ Is Nothing Then
        myDiscrep_.OldMenifestDocNo = curMenifest_.docNo
        myDiscrep_.OldMenifestNo = curMenifest_.MenifestNo
        myDiscrep_.menifestID = curMenifest_.ID
    End If
End Sub

Public Sub SetDataWasteData(selWaste As BWGWasteDataCR, Optional isAssigned As Boolean = True)
    Set curWasteData_ = selWaste
    With curWasteData_
        txtCode.Text = .WasteNo
        If isAssigned Then
            myDiscrep_.NewWasteDataID = .ID
            myDiscrep_.NewWasteName = .WasteName
            myDiscrep_.NewWasteNo = .WasteNo
        End If
    End With
End Sub

Public Sub SetDataTSDF(selTsdf As BWGTSDF, Optional isAssigned As Boolean = True)
    Set curTSDF_ = selTsdf
    With curTSDF_
        txtTsdf.Text = .CompanyName
        If isAssigned Then
            myDiscrep_.TSDFID = .ID
            myDiscrep_.TSDFName = .CompanyName
        End If
    End With
End Sub

Private Sub optAction_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            myDiscrep_.ActionTaken = "A"
        Case 1
            myDiscrep_.ActionTaken = "B"
        Case 2
            myDiscrep_.ActionTaken = "C"
    End Select
End Sub

Private Sub txtCode_Validate(Cancel As Boolean)
'    If Trim(txtCode.Text) = "" And Trim(curWasteData_.WasteNo) <> Trim(txtCode.Text) Then
'        Set curWasteData_ = Nothing
'        Set curWasteData_ = New BWGWasteDataCR
'        With myDiscrep_
'            .NewWasteDataID = ""
'            .NewWasteName = ""
'            .NewWasteNo = ""
'        End With
'    Else
'        If Trim(curWasteData_.WasteNo) <> Trim(txtCode.Text) Then
'            Set curWasteData_ = Nothing
'            Set curWasteData_ = New BWGWasteDataCR
'            With myDiscrep_
'                .NewWasteDataID = ""
'                .NewWasteName = ""
'                .NewWasteNo = ""
'            End With
'            Load frmWasteDataSearch
'            With frmWasteDataSearch
'                .getCriteria (txtCode.Text)
'                .setCallerForm Me
'                .Show 1
'            End With
'            If myDiscrep_.NewWasteDataID = "" Then txtCode.Text = ""
'        End If
'    End If
    myDiscrep_.NewWasteDataID = txtCode.Text
End Sub

Private Sub txtDesc_Validate(Cancel As Boolean)
    myDiscrep_.DiscrepDesc = Trim(txtDesc.Text)
End Sub

Private Sub txtDiscrepQty_Validate(Cancel As Boolean)
    If IsNumeric(txtDiscrepQty.Text) Then
        myDiscrep_.DiscrepQty = txtDiscrepQty.Text
    Else
        myDiscrep_.DiscrepQty = 0
    End If
End Sub

Private Sub txtRemDesc_Validate(Cancel As Boolean)
    myDiscrep_.ReasonAction = txtRemDesc.Text
End Sub

Private Sub txtReturnNo_Validate(Cancel As Boolean)
    myDiscrep_.ReturnMenifestID = txtReturnNo.Text
End Sub

Private Sub txtTsdf_Validate(Cancel As Boolean)
    If curTSDF_ Is Nothing Then Set curTSDF_ = New BWGTSDF
    If Trim(txtTsdf.Text) = "" And Trim(curTSDF_.CompanyName) <> Trim(txtTsdf.Text) Then
        Set curTSDF_ = Nothing
        Set curTSDF_ = New BWGTSDF
        With myDiscrep_
            .TSDFID = ""
            .TSDFName = ""
        End With
    Else
        If Trim(curTSDF_.CompanyName) <> Trim(txtTsdf.Text) Then
            Set curTSDF_ = Nothing
            Set curTSDF_ = New BWGTSDF
            With myDiscrep_
                .TSDFID = ""
                .TSDFName = ""
            End With
            Load frmTSDFSearch
            With frmTSDFSearch
                .getCriteria (txtTsdf.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If myDiscrep_.TSDFID = "" Then txtTsdf.Text = ""
        End If
    End If
End Sub

Private Sub PopulateUnit()
On Error Resume Next
Dim X%
'Dim tmpS As New BWGSearchManager
'    Set colUnit_ = tmpS.UnitGeneralSearch
    cboDiscrepUnit.Clear
    For X = 1 To colUnit_.Count
        cboDiscrepUnit.AddItem colUnit_(X).UnitName
    Next
'    Set tmpS = Nothing
End Sub
