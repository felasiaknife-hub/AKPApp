VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frUserProfile 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Picture         =   "frUserProfile.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5430
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   3390
      Picture         =   "frUserProfile.frx":0116
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5430
      Width           =   1185
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   4620
      Picture         =   "frUserProfile.frx":02D9
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5430
      Width           =   1185
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   5850
      Picture         =   "frUserProfile.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5430
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H8000000B&
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7080
      Picture         =   "frUserProfile.frx":0F60
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5430
      Width           =   1155
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
      TabIndex        =   21
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
      TabIndex        =   17
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
         TabIndex        =   18
         Top             =   1350
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2025
      Left            =   30
      TabIndex        =   26
      Top             =   3330
      Width           =   8265
      _ExtentX        =   14579
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ชื่อกลุ่มผู้ใช้ระบบ"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "คำอธิบาย"
         Object.Width           =   7937
      EndProperty
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
      Caption         =   "ชื่อย่อ"
      Height          =   195
      Index           =   6
      Left            =   6030
      TabIndex        =   25
      Top             =   1080
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อ"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   210
      TabIndex        =   20
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
      TabIndex        =   19
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
      Caption         =   "ข้อมูลส่วนบุคคลของ...."
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
      TabIndex        =   16
      Top             =   150
      Width           =   2265
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   0
      Picture         =   "frUserProfile.frx":14DC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10050
   End
End
Attribute VB_Name = "frUserProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myUser_ As AWSUser

Private Sub chkDisable_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    If chkDisable.Value = 0 Then
        myUser_.isActive = True
    Else
        myUser_.isActive = False
    End If
End Sub

Private Sub cmdClose_Click()
Dim myRes%
    If myUser_.isChanged Then
        myRes = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการเก็บข้อมูลก่อนปิดหรือไม่", vbYesNoCancel + vbQuestion)
        If myRes = vbYes Then
            cmdSave_Click
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
    Set myUser_ = New AWSUser
    With myUser_
        .ID = getNewProductID(modPrefix.UserPrefix)
    End With
    'chkAdmin.Enabled = True
    txtLoginName.Enabled = True
End Sub

Public Sub setCurrentUser(selUser As AWSUser)
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
'    If myUser_.isAdmin And frUserList.AdminCount = 1 Then
'        chkAdmin.Enabled = False
'    End If
End Sub

Private Sub cmdDel_Click()
    
    If MsgBox("ต้องการลบข้อมูลผู้ใช้คนนี้หรือไม่", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
'    If myUser_.isAdmin And frUserList.AdminCount = 1 Then
'        MsgBox "ไม่สามารถลบข้อมูลผู้ดูแลระบบคนสุดท้ายออกจากฐานข้อมูลได้", vbExclamation
'        Exit Sub
'    End If
    frUserList.DeleteUser myUser_.ID, False
    CreateNewUser
    ShowUserData
    txtFName.SetFocus
End Sub

Private Sub cmdNew_Click()
Dim myRes%
    If myUser_.isChanged Then
        myRes = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการเก็บข้อมูลก่อนสร้างใหม่หรือไม่", vbYesNoCancel + vbQuestion)
        If myRes = vbYes Then
            cmdSave_Click
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

Private Sub cmdSave_Click()
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
    myUser_.Password = Trim(txtPwd.Text)
    If myUser_.isChanged Then
        myUser_.Save
        myUser_.RefreshUserGroup
        If isFormLoaded("frUserList") Then frUserList.UpdateItem myUser_
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbInformation
    End If
End Sub

Private Sub Form_Load()
Dim x%
Dim tmpList As listItem
Dim tmpGroup As AWSUserGroup
    CreateNewUser
    PopulateUserGroup
    For x = 1 To ColUserGroup.Count
        Set tmpGroup = ColUserGroup(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpGroup.ID & "'", tmpGroup.Name)
        tmpList.SubItems(1) = tmpGroup.Description
        Set tmpList = Nothing
        Set tmpGroup = Nothing
    Next
End Sub

Private Sub ShowUserData()
    If myUser_ Is Nothing Then Exit Sub
    ClearScreen
    With myUser_
        txtFName.Text = .FirstName
        txtLName.Text = .LastName
        txtNickName.Text = .NickName
        txtAddr1.Text = .Address1
        txtAddr2.Text = .Address2
        txtAddr3.Text = .Address3
        txtPhone.Text = .PhoneNumber
        txtLoginName.Text = .LoginName
        txtPwd.Text = .Password
        txtConPwd.Text = .Password
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

Private Sub lvResult_ItemCheck(ByVal Item As MSComctlLib.listItem)
Dim tmpCol As Collection
    If myUser_ Is Nothing Then Exit Sub
    Set tmpCol = myUser_.getUserGroups
    If isExist(tmpCol, Item.Key) And Item.Checked = False Then
        tmpCol.Remove Item.Key
    ElseIf isExist(tmpCol, Item.Key) = False And Item.Checked = True Then
        tmpCol.Add ColUserGroup(Item.Key), Item.Key
    End If
    myUser_.isChanged = True
    Set tmpCol = Nothing
End Sub

Private Sub txtAddr1_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.Address1 = Trim(txtAddr1.Text)
End Sub

Private Sub txtAddr2_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.Address2 = Trim(txtAddr2.Text)
End Sub

Private Sub txtAddr3_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.Address3 = Trim(txtAddr3.Text)
End Sub

Private Sub txtFName_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.FirstName = Trim(txtFName.Text)
End Sub

Private Sub txtLName_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.LastName = Trim(txtLName.Text)
End Sub

Private Sub txtLoginName_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.LoginName = Trim(txtLoginName.Text)
End Sub

Private Sub txtNickName_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.NickName = Trim(txtNickName.Text)
End Sub

Private Sub txtPhone_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.PhoneNumber = Trim(txtPhone.Text)
End Sub

Private Sub txtPwd_Validate(Cancel As Boolean)
    If myUser_ Is Nothing Then Exit Sub
    myUser_.Password = Trim(txtPwd.Text)
End Sub

Private Sub ClearScreen()
Dim x%
    txtFName.Text = ""
    txtLName.Text = ""
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
End Sub
