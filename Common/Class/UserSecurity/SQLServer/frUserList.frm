VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frUserList 
   BackColor       =   &H80000009&
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   9855
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView lvResult 
      Height          =   3585
      Left            =   0
      TabIndex        =   0
      Top             =   930
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   6324
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ชื่อ-นามสกุล"
         Object.Width           =   7937
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "สถานะ"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายชื่อผู้ใช้ระบบ"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   270
      TabIndex        =   1
      Top             =   150
      Width           =   1605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   330
      X2              =   7140
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   240
      X2              =   7050
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   0
      Picture         =   "frUserList.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9690
   End
   Begin VB.Menu mnu_PopUser 
      Caption         =   "Menu"
      Begin VB.Menu mnu_AddUser 
         Caption         =   "เพิ่มข้อมูลผู้ใช้ระบบ"
      End
      Begin VB.Menu mnu_EditUser 
         Caption         =   "แก้ไขข้อมูลผู้ใช้ระบบ"
      End
      Begin VB.Menu mnu_DelUser 
         Caption         =   "ลบข้อมูลผู้ใช้ระบบ"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Close 
         Caption         =   "ปิด"
      End
   End
End
Attribute VB_Name = "frUserList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colUser_ As Collection
Dim adminCount_ As Integer

Private Sub Form_Load()
    mnu_PopUser.Visible = False
End Sub

Private Sub Form_Resize()
On Error Resume Next
    Image2.Width = Me.Width + 30
    With lvResult
        .Width = Me.Width - 165
        .Height = Me.Height - 1350
    End With
    
    LoadUserList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set colUser_ = Nothing
    adminCount_ = 0
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnu_PopUser, , x + Image2.Left, y + Image2.Top
    End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnu_PopUser, , x + Label1.Left, y + Label1.Top
    End If
End Sub

Private Sub lvResult_DblClick()
    EditUser
End Sub

Private Sub lvResult_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpList As listItem
    If KeyCode = 46 Then
        Set tmpList = lvResult.SelectedItem
        If tmpList Is Nothing Then
            MsgBox "กรุณาเลือกรายการที่ต้องการลบข้อมูล", vbExclamation
            Exit Sub
        End If
        frUserList.DeleteUser Replace(tmpList.Key, "'", "")
        Set tmpList = Nothing
    End If
End Sub

Private Sub lvResult_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnu_PopUser, , x + lvResult.Left, y + lvResult.Top
    End If
End Sub

Private Sub mnu_AddUser_Click()
    frUserProfile.Show 1
End Sub

Private Sub LoadUserList()
Dim x%
Dim tmpList As listItem
Dim tmpUser As AWSUser
'Dim tmpsearch As New MMSSearchManager
    Set colUser_ = modUserSecurity.UserSearch
    lvResult.ListItems.Clear
    For x = 1 To colUser_.Count
        Set tmpUser = colUser_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpUser.ID & "'", tmpUser.FirstName & " " & tmpUser.LastName)
        If tmpUser.isActive Then
            tmpList.SubItems(1) = "พร้อมใช้งาน"
        Else
            tmpList.SubItems(1) = "ระงับการใช้งาน"
        End If
'        If tmpUser.isAdmin Then
'            adminCount_ = adminCount_ + 1
'        End If
        Set tmpList = Nothing
        Set tmpUser = Nothing
    Next
'    Set tmpsearch = Nothing
End Sub

Public Property Get AdminCount() As Integer
    AdminCount = adminCount_
End Property
Public Property Let AdminCount(ByVal newVal As Integer)
    adminCount_ = newVal
End Property

Public Sub DeleteUser(selID As String, Optional showMsg As Boolean = True)
Dim tmpUser As AWSUser
Dim tmpKey As String
    If colUser_ Is Nothing Then Exit Sub
    tmpKey = Trim("'" & selID & "'")
    If isExist(colUser_, tmpKey) Then
        Set tmpUser = colUser_(tmpKey)
        If showMsg Then
            If MsgBox("ต้องการลบข้อมูลผู้ใช้คนนี้หรือไม่", vbQuestion + vbYesNo) = vbNo Then
                Exit Sub
            End If
'            If tmpUser.isAdmin And frUserList.AdminCount = 1 Then
'                MsgBox "ไม่สามารถลบข้อมูลผู้ดูแลระบบคนสุดท้ายออกจากฐานข้อมูลได้", vbExclamation
'                Exit Sub
'            End If
        End If
        tmpUser.Delete
'        If tmpUser.isAdmin Then
'            adminCount_ = adminCount_ - 1
'        End If
        colUser_.Remove tmpKey
        lvResult.ListItems.Remove tmpKey
        Set tmpUser = Nothing
    End If
End Sub

Public Function getUserList() As Collection
    Set getUserList = colUser_
End Function

Public Sub SaveUser(selUser As AWSUser)
Dim tmpKey As String
Dim tmpList As listItem
    tmpKey = Trim("'" & selUser.ID & "'")
    If isExist(colUser_, tmpKey) Then
        colUser_.Remove tmpKey
        Set tmpList = lvResult.ListItems(tmpKey)
        tmpList.Text = selUser.FirstName & "  " & selUser.LastName
    Else
        Set tmpList = lvResult.ListItems.Add(, tmpKey, selUser.FirstName & "  " & selUser.LastName)
    End If
    If selUser.isActive Then
        tmpList.SubItems(1) = "พร้อมใช้งาน"
    Else
        tmpList.SubItems(1) = "ระงับการใช้งาน"
    End If
    Set tmpList = Nothing

    colUser_.Add selUser, tmpKey
'    If selUser.isAdmin Then
'        adminCount_ = adminCount_ + 1
'    End If

End Sub

Public Sub EditUser()
Dim tmpUser As AWSUser
    If frUserList.lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกรายการที่ต้องการแก้ไขข้อมูล", vbExclamation
        Exit Sub
    End If
    Set tmpUser = frUserList.getUserList(frUserList.lvResult.SelectedItem.Key)
    Load frUserProfile
    With frUserProfile
        .setCurrentUser tmpUser.DuplicateUser
        Set tmpUser = Nothing
        .Show 1
    End With
End Sub

Public Sub UpdateItem(selItem As AWSUser)
Dim tmpKey As String
Dim tmpList As listItem
    With selItem
        tmpKey = Trim("'" & .ID & "'")
        If isExist(colUser_, tmpKey) Then
            Set tmpList = lvResult.ListItems(tmpKey)
            tmpList.Text = .FirstName & " " & .LastName
            colUser_.Remove tmpKey
        Else
            Set tmpList = lvResult.ListItems.Add(, "'" & .ID & "'", .FirstName & " " & .LastName)
        End If
        If .isActive Then
            tmpList.SubItems(1) = "พร้อมใช้งาน"
        Else
            tmpList.SubItems(1) = "ระงับการใช้งาน"
        End If
        colUser_.Add selItem, tmpKey
    End With
    Set tmpList = Nothing
End Sub

Private Sub mnu_Close_Click()
    Unload Me
End Sub

Private Sub mnu_DelUser_Click()
Dim tmpList As listItem
    Set tmpList = lvResult.SelectedItem
    If tmpList Is Nothing Then
        MsgBox "กรุณาเลือกรายการที่ต้องการลบข้อมูล", vbExclamation
        Exit Sub
    End If
    frUserList.DeleteUser Replace(tmpList.Key, "'", "")
    Set tmpList = Nothing
End Sub

Private Sub mnu_EditUser_Click()
    EditUser
End Sub
