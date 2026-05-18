VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserGroup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F140"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkRole 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "สามารถเลือกข้อมูลได้"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   20
      Top             =   3840
      Width           =   2085
   End
   Begin VB.CheckBox chkRole 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "สามารถอนุมัติข้อมูลได้"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   5760
      TabIndex        =   9
      Top             =   3510
      Width           =   2085
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4680
      Width           =   945
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "ลบข้อมูล"
      Height          =   795
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4680
      Width           =   945
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4680
      Width           =   945
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4680
      Width           =   945
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4680
      Width           =   945
   End
   Begin VB.CommandButton cmdNone 
      Caption         =   "-"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4260
      Width           =   495
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "+"
      Enabled         =   0   'False
      Height          =   285
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4260
      Width           =   495
   End
   Begin VB.CheckBox chkRole 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "สามารถพิมพ์ข้อมูลได้"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   8
      Top             =   3210
      Width           =   2085
   End
   Begin VB.CheckBox chkRole 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "สามารถลบข้อมูลได้"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   5760
      TabIndex        =   7
      Top             =   2910
      Width           =   2085
   End
   Begin VB.CheckBox chkRole 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "สามารถแก้ไขข้อมูลมูลเก่าได้"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   6
      Top             =   2610
      Width           =   2295
   End
   Begin VB.CheckBox chkRole 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "สามารถเพิ่มข้อมูลใหม่ได้"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   5
      Top             =   2310
      Width           =   2085
   End
   Begin VB.CheckBox chkRole 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "สามารถเรียกดูข้อมูลได้"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   5760
      TabIndex        =   4
      Top             =   1980
      Width           =   2085
   End
   Begin VB.CheckBox chkDisable 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ระงับการใช้งานชั่วคราว"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3540
      TabIndex        =   1
      Top             =   1050
      Width           =   2145
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1020
      TabIndex        =   2
      Top             =   1410
      Width           =   6825
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1020
      TabIndex        =   0
      Top             =   1050
      Width           =   2235
   End
   Begin MSComctlLib.ListView lvFormList 
      Height          =   2715
      Left            =   30
      TabIndex        =   3
      Top             =   1890
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ระบบงาน"
         Object.Width           =   8996
      EndProperty
   End
   Begin VB.Shape Shape1 
      Height          =   2715
      Index           =   1
      Left            =   5640
      Top             =   1890
      Width           =   2535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำอธิบาย"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   19
      Top             =   1440
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อกลุ่ม"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   18
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ข้อมูลกลุ่มผู้ใช้ระบบ"
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
      TabIndex        =   17
      Top             =   150
      Width           =   1950
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   150
      X2              =   6960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   -30
      Picture         =   "frUserGroup.frx":0000
      Stretch         =   -1  'True
      Top             =   -30
      Width           =   10050
   End
End
Attribute VB_Name = "frmUserGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myGroup_ As AWSUserGroup

Private Sub chkDisable_Click()
    If Me.Visible Then
        If Me.ActiveControl.Name = "chkDisable" Then
            If myGroup_ Is Nothing Then Exit Sub
            If chkDisable.Value = 0 Then
                myGroup_.isActive = True
            Else
                myGroup_.isActive = False
            End If
        End If
    End If
End Sub

Private Sub chkRole_Click(Index As Integer)
    If Me.Visible = False Then Exit Sub
    If Me.ActiveControl.Name <> "chkRole" Then Exit Sub
    SetAccessRoleValue Index
End Sub

Private Sub cmdAll_Click()
Dim x%
    For x = 0 To chkRole.UBound
        chkRole(x).Value = 1
        SetAccessRoleValue x
    Next
End Sub

Private Sub cmdClear_Click()
Dim selVal As Integer
    If myGroup_.isChanged Then
        selVal = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนการสร้างรายการใหม่หรือไม่?", vbQuestion + vbYesNoCancel)
        If selVal = vbYes Then
            cmdSave_Click
            If myGroup_.isChanged Then Exit Sub
        ElseIf selVal = vbCancel Then
            Exit Sub
        End If
    End If
    ClearScreen
    CreateNewGroup
    txtName.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
Dim tmpKey As String
Dim tmpStr As String
    If myGroup_.isSystem Then
        MsgBox "ไม่สามารถลบข้อมูลกลุ่มผู้ใช้ระบบนี้ได้", vbExclamation
        Exit Sub
    End If
    If MsgBox("ต้องการลบข้อมูลกลุ่มผู้ใช้ระบบนี้หรือไม่?", vbYesNo) = vbYes Then
        tmpKey = "'" & myGroup_.ID & "'"
        tmpStr = myGroup_.Delete
        If Trim(tmpStr) = "" Then
            If isExist(ColUserGroup, tmpKey) Then
                ColUserGroup.Remove tmpKey
            End If
            ClearScreen
            CreateNewGroup
            txtName.SetFocus
        Else
            MsgBox tmpStr, vbExclamation
        End If
    End If
End Sub

Private Sub cmdNone_Click()
Dim x%
    For x = 0 To chkRole.UBound
        chkRole(x).Value = 0
        SetAccessRoleValue x
    Next
End Sub

Private Sub cmdSave_Click()
    If Trim(txtName.Text) = "" Then
        MsgBox "กรุณาระบุชื่อกลุ่มผู้ใช้ ก่อนจัดเก็บข้อมูล", vbExclamation
        txtName.SetFocus
        Exit Sub
    End If
    If myGroup_.getAccessRoles.Count = 0 Then
        MsgBox "กรุณาเลือกระบบงานและกำหนดขอบเขตการใช้งายอย่างน้อย 1 รายการ", vbExclamation
        lvFormList.SetFocus
        Exit Sub
    End If
    If myGroup_.isChanged Then
        myGroup_.Save
        If isExist(ColUserGroup, "'" & myGroup_.ID & "'") = False Then
            ColUserGroup.Add myGroup_, "'" & myGroup_.ID & "'"
        End If
    End If
    If myGroup_.isChanged = False Then
        MsgBox "จัดเก็บข้อมูลเรียบร้อยแล้ว", vbExclamation
    End If
End Sub

Private Sub cmdSearch_Click()
Dim selResult As Integer
    If myGroup_.isChanged Then
        selResult = MsgBox("ข้อมูลมีการเปลี่ยนแปลงต้องการจัดเก็บข้อมูลก่อนทำงานต่อหรือไม่?", vbQuestion + vbYesNoCancel)
        If selResult = vbYes Then
            cmdSave_Click
            If myGroup_.isChanged Then Exit Sub
        ElseIf selResult = vbCancel Then
            Exit Sub
        End If
    End If

    Load frmUserGroupSearch
    With frmUserGroupSearch
        .SetCallerForm Me
        .Show 1
    End With
End Sub

Private Sub Form_Load()
Dim x%
Dim tmpArr
Dim tmpList As ListItem
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    PopulateFormList
    PopulateUserGroup
    For x = 1 To ColFormList.Count
        tmpArr = Split(ColFormList(x), "_")
        lvFormList.ListItems.Add , "'" & tmpArr(0) & "'", tmpArr(1)
    Next
    CreateNewGroup
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdClear.Picture = .picNew.Picture
        cmdSearch.Picture = .picSearch.Picture
        cmdClose.Picture = .picExit.Picture
        cmdDel.Picture = .picDelete.Picture
    End With

End Sub

Private Sub CreateNewGroup()
    Set myGroup_ = Nothing
    Set myGroup_ = New AWSUserGroup
    With myGroup_
        '.ID = getNewProductID(modPrefix.UserGroupPrefix)
        .ID = GetGUID
        .isActive = True
        .isChanged = False
    End With
End Sub

Private Sub ClearScreen()
    txtName.Text = ""
    chkDisable.Value = 0
    txtDesc.Text = ""
    chkRole(0).Value = 0
    chkRole(1).Value = 0
    chkRole(2).Value = 0
    chkRole(3).Value = 0
    chkRole(4).Value = 0
    chkRole(5).Value = 0
    chkRole(6).Value = 0
    SetRoleEnabled False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim selVal As Integer
    If myGroup_.isChanged Then
        selVal = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนปิดหน้าจอนี้หรือไม่?", vbQuestion + vbYesNoCancel)
        If selVal = vbYes Then
            cmdSave_Click
            If myGroup_.isChanged Then Cancel = 1
        ElseIf selVal = vbCancel Then
            Cancel = 1
        End If
    End If
End Sub

Private Sub lvFormList_ItemClick(ByVal item As MSComctlLib.ListItem)
Dim tmpKey As String
Dim tmpCol As Collection
Dim tmpRole As AWSAccessRole
    tmpKey = Trim(item.Key)
    Set tmpCol = myGroup_.getAccessRoles
    SetRoleEnabled True
    chkRole(0).Value = 0
    chkRole(1).Value = 0
    chkRole(2).Value = 0
    chkRole(3).Value = 0
    chkRole(4).Value = 0
    chkRole(5).Value = 0
    chkRole(6).Value = 0
    If isExist(tmpCol, tmpKey) Then
        Set tmpRole = tmpCol(tmpKey)
        If tmpRole.CanAccess Then chkRole(0).Value = 1
        If tmpRole.CanAddNew Then chkRole(1).Value = 1
        If tmpRole.CanUpdate Then chkRole(2).Value = 1
        If tmpRole.CanDelete Then chkRole(3).Value = 1
        If tmpRole.CanPrint Then chkRole(4).Value = 1
        If tmpRole.CanApprove Then chkRole(5).Value = 1
        If tmpRole.CanSelect Then chkRole(6).Value = 1
        Set tmpRole = Nothing
    End If
    Set tmpCol = Nothing
End Sub

Public Sub SetData(selGroup As AWSUserGroup)
    ClearScreen
    Set myGroup_ = selGroup
    If myGroup_ Is Nothing Then Exit Sub

    txtName.Text = myGroup_.Name
    txtDesc.Text = myGroup_.Description
    If myGroup_.isActive Then
        chkDisable.Value = 0
    Else
        chkDisable.Value = 1
    End If

End Sub

Private Sub txtDesc_Validate(Cancel As Boolean)
    If myGroup_ Is Nothing Then Exit Sub
    myGroup_.Description = Trim(txtDesc.Text)
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    If myGroup_ Is Nothing Then Exit Sub
    myGroup_.Name = Trim(txtName.Text)
End Sub

Private Sub SetRoleEnabled(selVal As Boolean)
Dim x%
    For x = 0 To chkRole.UBound
        chkRole(x).Enabled = selVal
    Next
    cmdAll.Enabled = selVal
    cmdNone.Enabled = selVal
End Sub

Private Sub SetAccessRoleValue(Index As Integer)
Dim tmpRole As AWSAccessRole
Dim tmpCol As Collection
Dim tmpKey As String
Dim tmpVal As Boolean
    If myGroup_ Is Nothing Then Exit Sub
    If lvFormList.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกระบบงานที่ต้องการ", vbExclamation
        lvFormList.SetFocus
        Exit Sub
    End If
    Set tmpCol = myGroup_.getAccessRoles
    tmpKey = lvFormList.SelectedItem.Key
    If isExist(tmpCol, tmpKey) Then
        Set tmpRole = tmpCol(tmpKey)
    Else
        Set tmpRole = New AWSAccessRole
        'tmpRole.ID = getNewProductID(modPrefix.RoleAccessPrefix) & index
        tmpRole.GroupID = myGroup_.ID
        tmpRole.FormName = Replace(tmpKey, "'", "")
        tmpCol.Add tmpRole, "'" & tmpRole.FormName & "'"
    End If
    If chkRole(Index).Value = 1 Then
        tmpVal = True
    Else
        tmpVal = False
    End If
    Select Case Index
        Case 0
            tmpRole.CanAccess = tmpVal
        Case 1
            tmpRole.CanAddNew = tmpVal
        Case 2
            tmpRole.CanUpdate = tmpVal
        Case 3
            tmpRole.CanDelete = tmpVal
        Case 4
            tmpRole.CanPrint = tmpVal
        Case 5
            tmpRole.CanApprove = tmpVal
        Case 6
            tmpRole.CanSelect = tmpVal
    End Select
    If tmpRole.isChanged Then myGroup_.isChanged = True
    Set tmpRole = Nothing
    Set tmpCol = Nothing
End Sub
