VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserProfile 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F143"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtDept 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5145
      TabIndex        =   32
      Top             =   3420
      Width           =   2655
   End
   Begin VB.TextBox txtPosition 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1170
      TabIndex        =   11
      Top             =   3420
      Width           =   2655
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   4350
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6450
      Width           =   945
   End
   Begin VB.CommandButton cmdChangePwd 
      Caption         =   "เปลี่ยนรหัสผ่าน"
      Height          =   795
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6450
      Width           =   1455
   End
   Begin VB.CheckBox chkDisable 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ระงับการใช้งานชั่วคราว"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4260
      TabIndex        =   10
      Top             =   2910
      Width           =   2145
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   "สิทธิการใช้โปรแกรม"
      Height          =   795
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7050
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6450
      Width           =   945
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   5340
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6450
      Width           =   945
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6450
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000B&
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6450
      Width           =   915
   End
   Begin VB.TextBox txtNickName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6510
      TabIndex        =   2
      Top             =   1050
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ข้อมูลการเข้าใช้ระบบ"
      ForeColor       =   &H80000008&
      Height          =   1425
      Index           =   1
      Left            =   4230
      TabIndex        =   26
      Top             =   1410
      Width           =   3885
      Begin VB.TextBox txtConPwd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   960
         Width           =   1965
      End
      Begin VB.TextBox txtPwd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   630
         Width           =   1965
      End
      Begin VB.TextBox txtLoginName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Top             =   300
         Width           =   1965
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยืนยันรหัสผ่าน"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   540
         TabIndex        =   29
         Top             =   990
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสผ่าน"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   5
         Left            =   990
         TabIndex        =   28
         Top             =   660
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อในการเข้าระบบ"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   27
         Top             =   330
         Width           =   1290
      End
   End
   Begin VB.TextBox txtFName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Top             =   1050
      Width           =   2235
   End
   Begin VB.TextBox txtLName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Top             =   1050
      Width           =   2235
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ที่อยู่"
      ForeColor       =   &H80000008&
      Height          =   1785
      Index           =   0
      Left            =   180
      TabIndex        =   22
      Top             =   1410
      Width           =   3885
      Begin VB.TextBox txtAddr1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   3
         Top             =   300
         Width           =   3345
      End
      Begin VB.TextBox txtAddr2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   4
         Top             =   630
         Width           =   3345
      End
      Begin VB.TextBox txtAddr3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   300
         TabIndex        =   5
         Top             =   960
         Width           =   3345
      End
      Begin VB.TextBox txtPhone 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   990
         TabIndex        =   6
         Top             =   1290
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "โทรศัพท์"
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   23
         Top             =   1350
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   30
      TabIndex        =   12
      Top             =   3780
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   4577
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ชื่อกลุ่มผู้ใช้ระบบ"
         Object.Width           =   5292
      EndProperty
   End
   Begin MSComctlLib.ListView lvCompany 
      Height          =   2580
      Left            =   4200
      TabIndex        =   13
      Top             =   3795
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   4551
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "บริษัทที่สังกัด"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "แผนก"
      Height          =   195
      Index           =   8
      Left            =   4455
      TabIndex        =   33
      Top             =   3480
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ตำแหน่ง"
      Height          =   195
      Index           =   7
      Left            =   480
      TabIndex        =   31
      Top             =   3480
      Width           =   600
   End
   Begin VB.Shape Shape1 
      Height          =   2415
      Left            =   30
      Top             =   930
      Width           =   8265
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัส"
      Height          =   195
      Index           =   6
      Left            =   6120
      TabIndex        =   30
      Top             =   1080
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อ"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   25
      Top             =   1080
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "นามสกุล"
      Height          =   195
      Index           =   1
      Left            =   2850
      TabIndex        =   24
      Top             =   1080
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   150
      X2              =   6960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   7050
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ข้อมูลผู้ใช้ระบบ"
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
      Left            =   180
      TabIndex        =   21
      Top             =   150
      Width           =   1530
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   0
      Picture         =   "frUserProfile.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10050
   End
End
Attribute VB_Name = "frmUserProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myUser_ As BWGCompanyStaff
Dim colCompany_ As Collection

Private Sub chkDisable_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    If chkDisable.Value = 0 Then
        myUser_.isActive = True
    Else
        myUser_.isActive = False
    End If
End Sub

Private Sub cmdChangePwd_Click()
    Load frmChangePwd
    With frmChangePwd
        .SetShowUser myUser_
        .Show 1
    End With
End Sub

Private Sub cmdClose_Click()
Dim myRes%
    If myUser_.isChanged Then
        myRes = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการเก็บข้อมูลก่อนปิดหรือไม่", vbYesNoCancel + vbQuestion)
        If myRes = vbYes Then
            CmdSave_Click
            If myUser_.isChanged Then
                Exit Sub
            End If
        ElseIf myRes = vbCancel Then
            Exit Sub
        End If
    End If

    Unload Me
End Sub

Private Sub CreateNewUser()
    Set myUser_ = Nothing
    Set myUser_ = New BWGCompanyStaff
    With myUser_
        '.ID = getNewProductID(modPrefix.UserPrefix)
        .ID = GetGUID
    End With
    'chkAdmin.Enabled = True
    txtLoginName.Enabled = True
End Sub

Public Sub setCurrentUser(selUser As BWGCompanyStaff)
Dim x%
Dim tmpCol As Collection
Dim tmpGroup As AWSUserGroup
Dim tmpKey As String
    Set myUser_ = Nothing
    Set myUser_ = selUser
    ShowUserData
    txtLoginName.Enabled = False
    Set tmpCol = myUser_.getUserGroups
    For x = 1 To tmpCol.Count
        tmpKey = Trim("'" & tmpCol(x).ID & "'")
        If isExist(lvResult.ListItems, tmpKey) Then
            lvResult.ListItems(tmpKey).Checked = True
        Else
            lvResult.ListItems(tmpKey).Checked = False
        End If
    Next

    Set tmpCol = myUser_.WorkCompanies
    For x = 1 To tmpCol.Count
        tmpKey = Trim("'" & tmpCol(x).ID & "'")
        If isExist(lvCompany.ListItems, tmpKey) Then
            lvCompany.ListItems(tmpKey).Checked = True
        Else
            lvCompany.ListItems(tmpKey).Checked = False
        End If
    Next
    Set tmpCol = Nothing
'    If myUser_.isAdmin And frmUserList.AdminCount = 1 Then
'        chkAdmin.Enabled = False
'    End If
End Sub

Private Sub cmdDel_Click()
Dim tmpStr As String
    If MsgBox("ต้องการลบข้อมูลผู้ใช้คนนี้หรือไม่", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
'    If myUser_.isAdmin And frmUserList.AdminCount = 1 Then
'        MsgBox "ไม่สามารถลบข้อมูลผู้ดูแลระบบคนสุดท้ายออกจากฐานข้อมูลได้", vbExclamation
'        Exit Sub
'    End If
    'frmUserList.DeleteUser myUser_.ID, False
    tmpStr = myUser_.Delete
    If Trim(tmpStr) = "" Then
        CreateNewUser
        ShowUserData
        txtFName.SetFocus
    Else
        MsgBox tmpStr, vbExclamation
    End If
End Sub

Private Sub cmdNew_Click()
Dim myRes%
    If myUser_.isChanged Then
        myRes = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการเก็บข้อมูลก่อนสร้างใหม่หรือไม่", vbYesNoCancel + vbQuestion)
        If myRes = vbYes Then
            CmdSave_Click
            If myUser_.isChanged Then
                Exit Sub
            End If
        ElseIf myRes = vbCancel Then
            Exit Sub
        End If
    End If
    CreateNewUser
    ShowUserData
    txtFName.SetFocus
End Sub

Private Sub CmdSave_Click()
Dim isSelected As Boolean
Dim x%
    If Trim(txtFName.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อพนักงาน", vbExclamation
        txtFName.SetFocus
        Exit Sub
    End If
    If Trim(txtLoginName.Text) = "" Then
        MsgBox "กรุณาใส่ชื่อที่ใช้ในการเข้าระบบ", vbExclamation
        txtLoginName.SetFocus
        Exit Sub
    End If
    If Trim(txtPwd.Text) = "" Then
        MsgBox "กรุณาใส่รหัสผ่าน", vbExclamation
        txtPwd.SetFocus
        Exit Sub
    End If
    If Trim(txtConPwd.Text) = "" Then
        MsgBox "กรุณายืนยันรหัสผ่าน", vbExclamation
        txtConPwd.SetFocus
        Exit Sub
    End If
    If Trim(txtPwd.Text) <> Trim(txtConPwd.Text) Then
        MsgBox "การยืนยันรหัสผ่านไม่ตรงกับรหัสผ่าน กรุณาตรวจสอบ", vbExclamation
        txtConPwd.SetFocus
        Exit Sub
    End If
    isSelected = False
    For x = 1 To lvResult.ListItems.Count
        If lvResult.ListItems(x).Checked Then
            isSelected = True
            Exit For
        End If
    Next
    If isSelected = False Then
        MsgBox "กรุณาเลือกชื่อกลุ่มผู้ใช้ระบบอย่างน้อย 1 กลุ่ม", vbExclamation
        lvResult.SetFocus
        Exit Sub
    End If
    isSelected = False
    For x = 1 To lvCompany.ListItems.Count
        If lvCompany.ListItems(x).Checked Then
            isSelected = True
            Exit For
        End If
    Next
    If isSelected = False Then
        MsgBox "กรุณาระบุบริษัทต้นสังกัดของพนักงานคนนี้อย่างน้อย 1 บริษัท", vbExclamation
        lvCompany.SetFocus
        Exit Sub
    End If
    myUser_.LoginPwd = Trim(txtPwd.Text)
    If myUser_.isChanged Then
        If myUser_.Save = "" Then MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
        myUser_.RefreshUserGroup
        If isFormLoaded("frmUserList") Then
            frmUserList.UpdateItem myUser_
         End If
    End If
End Sub

Private Sub cmdSearch_Click()
Dim myRes%
    If myUser_.isChanged Then
        myRes = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการเก็บข้อมูลก่อนสร้างใหม่หรือไม่", vbYesNoCancel + vbQuestion)
        If myRes = vbYes Then
            CmdSave_Click
            If myUser_.isChanged Then
                Exit Sub
            End If
        ElseIf myRes = vbCancel Then
            Exit Sub
        End If
    End If
    Load frmUserProfileSearch
    With frmUserProfileSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim x%
Dim tmpList As ListItem
Dim tmpGroup As AWSUserGroup
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdNew.Picture = .picNew.Picture
        cmdDel.Picture = .picDelete.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    CreateNewUser
    PopulateUserGroup
    PopulateCompany
    For x = 1 To ColUserGroup.Count
        Set tmpGroup = ColUserGroup(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpGroup.ID & "'", tmpGroup.Name)
        'tmpList.SubItems(1) = tmpGroup.Description
        Set tmpList = Nothing
        Set tmpGroup = Nothing
    Next
End Sub

Private Sub ShowUserData()
Dim EncObj As New RC4Crypt
    If myUser_ Is Nothing Then Exit Sub
    ClearScreen
    With myUser_
        txtFName.Text = .StaffFName
        txtLName.Text = .StaffLName
        txtNickName.Text = .StaffNickName
        txtAddr1.Text = .Addr1
        txtAddr2.Text = .Addr2
        txtAddr3.Text = .Addr3
        txtPhone.Text = .TelNo
        txtPosition.Text = .PositionName
        txtLoginName.Text = .LoginName
        txtPwd.Text = EncObj.EncryptString(Trim(.LoginPwd), ENC_KEY)
        txtConPwd.Text = txtPwd.Text
        TxtDept.Text = .Dept
        If .isActive Then
            chkDisable.Value = 0
        Else
            chkDisable.Value = 1
        End If
'        If .isAdmin Then chkAdmin.Value = 1
'        If .isTechnician Then chkTech.Value = 1
'        If .isReception Then chkRecp.Value = 1
'        If .isSale Then chkSale.Value = 1
    End With
End Sub

Private Sub lvCompany_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim tmpCol As Collection
    If myUser_ Is Nothing Then Exit Sub
    Set tmpCol = myUser_.WorkCompanies
    If isExist(tmpCol, item.Key) And item.Checked = False Then
        tmpCol.Remove item.Key
    ElseIf isExist(tmpCol, item.Key) = False And item.Checked = True Then
        tmpCol.Add colCompany_(item.Key), item.Key
    End If
    myUser_.isChanged = True
    Set tmpCol = Nothing
End Sub

Private Sub lvResult_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim tmpCol As Collection
    If myUser_ Is Nothing Then Exit Sub
    Set tmpCol = myUser_.getUserGroups
    If isExist(tmpCol, item.Key) And item.Checked = False Then
        tmpCol.Remove item.Key
    ElseIf isExist(tmpCol, item.Key) = False And item.Checked = True Then
        tmpCol.Add ColUserGroup(item.Key), item.Key
    End If
    myUser_.isChanged = True
    Set tmpCol = Nothing
End Sub

Private Sub txtAddr1_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.Addr1 = Trim(txtAddr1.Text)
End Sub

Private Sub txtAddr2_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.Addr2 = Trim(txtAddr2.Text)
End Sub

Private Sub txtAddr3_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.Addr3 = Trim(txtAddr3.Text)
End Sub

Private Sub TxtDept_Validate(Cancel As Boolean)
    myUser_.Dept = TxtDept.Text
End Sub

Private Sub txtFName_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.StaffFName = Trim(txtFName.Text)
End Sub

Private Sub txtLName_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.StaffLName = Trim(txtLName.Text)
End Sub

Private Sub txtLoginName_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.LoginName = Trim(txtLoginName.Text)
End Sub

Private Sub txtNickName_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.StaffNickName = Trim(txtNickName.Text)
    myUser_.StaffCode = Trim(txtNickName.Text)
End Sub

Private Sub txtPhone_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.TelNo = Trim(txtPhone.Text)
End Sub

Private Sub txtPosition_Change()
    myUser_.PositionName = txtPosition.Text
End Sub

Private Sub txtPwd_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.LoginPwd = Trim(txtPwd.Text)
End Sub

Private Sub ClearScreen()
Dim x%
    txtFName.Text = ""
    txtLName.Text = ""
    txtPosition.Text = ""
    txtNickName.Text = ""
    txtAddr1.Text = ""
    txtAddr2.Text = ""
    txtAddr3.Text = ""
    txtPhone.Text = ""
    txtLoginName.Text = ""
    txtPwd.Text = ""
    txtConPwd.Text = ""
    chkDisable.Value = 0
'    chkAdmin.Value = 0
'    chkTech.Value = 0
'    chkRecp.Value = 0
'    chkSale.Value = 0
    For x = 1 To lvResult.ListItems.Count
        lvResult.ListItems(x).Checked = False
    Next
    For x = 1 To lvCompany.ListItems.Count
        lvCompany.ListItems(x).Checked = False
    Next

End Sub

Public Sub DisableFormPart()
    txtLoginName.Locked = True
    txtPwd.Locked = True
    txtConPwd.Locked = True
    chkDisable.Enabled = False
    lvResult.Enabled = False
    lvCompany.Enabled = False
    cmdNew.Enabled = False
    cmdDel.Enabled = False
    cmdChangePwd.Visible = True
End Sub

Private Sub PopulateCompany()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colCompany_ = tmpS.CompanySearch
    Set tmpS = Nothing
    lvCompany.ListItems.Clear
    For x = 1 To colCompany_.Count
        lvCompany.ListItems.Add , "'" & colCompany_(x).ID & "'", colCompany_(x).CompanyNameTH
    Next
End Sub

