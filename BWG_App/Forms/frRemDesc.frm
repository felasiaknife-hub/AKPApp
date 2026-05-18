VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRemDesc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F113"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "ตกลง"
      Height          =   795
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5190
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5190
      Width           =   1035
   End
   Begin VB.TextBox txtShortName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1770
      MaxLength       =   1000
      TabIndex        =   0
      Top             =   870
      Width           =   4665
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5190
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5775
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5190
      Width           =   1035
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3825
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   6747
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
         Text            =   "ชื่อหน่วย"
         Object.Width           =   11994
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   5640
      Y1              =   405
      Y2              =   405
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   90
      X2              =   5790
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รายการหมายเหตุสำหรับใบเสนอราคา"
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
      Left            =   225
      TabIndex        =   6
      Top             =   0
      Width           =   3180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "รายละเอียดหมายเหตุ"
      Height          =   195
      Index           =   8
      Left            =   210
      TabIndex        =   5
      Top             =   930
      Width           =   1485
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frRemDesc.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6870
   End
End
Attribute VB_Name = "frmRemDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colRemDesc_ As Collection
Dim curRemDesc_ As BWGRemarkList
Dim curRole_ As AWSRoleType
Dim callerForm_ As Form

Public Sub SetCallerForm(frm As Form, Optional setYear As Integer = 0)
    Set callerForm_ = frm
End Sub

Private Sub cmdClose_Click()
Dim aa
    If Trim(txtShortName.Text) <> Trim(curRemDesc_.Remdesc) Then
        aa = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนปิดหรือไม่", vbQuestion + vbYesNoCancel)
        If aa = vbYes Then
            CmdSave_Click
            If Trim(txtShortName.Text) <> Trim(curRemDesc_.Remdesc) Then
                Exit Sub
            End If
        ElseIf aa = vbCancel Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdNew_Click()
Dim aa
    If Trim(txtShortName.Text) <> Trim(curRemDesc_.Remdesc) Then
        aa = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่", vbQuestion + vbYesNoCancel)
        If aa = vbYes Then
            CmdSave_Click
            If Trim(txtShortName.Text) <> Trim(curRemDesc_.Remdesc) Then
                Exit Sub
            End If
        ElseIf aa = vbCancel Then
            Exit Sub
        End If
    End If
    CreateNewRemDesc
    ClearScreen
    txtShortName.SetFocus
End Sub

Private Sub cmdOK_Click()
Dim x%
    For x = 1 To lvList.ListItems.Count
        If lvList.ListItems(x).Checked Then
            callerForm_.SetDataRemark colRemDesc_(lvList.ListItems(x).Key)
        End If
    Next
    Unload Me
End Sub

Private Sub CmdSave_Click()
Dim sqlStr As String
Dim tmpList As ListItem
Dim tmpCon As ADODB.Connection
Dim tmpStr As String
    'tmpStr = CurrentUser.VerifyAccessRole(Me.Name, curRole_)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If

    If curRemDesc_ Is Nothing Then Exit Sub
    If Trim(txtShortName.Text) = vbNullString Then
        MsgBox "กรุณาใส่รายละเอียดของหมายเหตุก่อนเก็บข้อมูล", vbExclamation
        txtShortName.SetFocus
        Exit Sub
    End If

    If Trim(txtShortName.Text) <> Trim(curRemDesc_.Remdesc) Then
        If Trim(curRemDesc_.ID) = "" Then
            curRemDesc_.ID = GetGUID
            DBConnExc "Insert into tbRemarkList (RemarkID,RemDesc) Values ('" & Trim(curRemDesc_.ID) & "','" & (txtShortName.Text) & "')"
            Set tmpList = lvList.ListItems.Add(, "'" & curRemDesc_.ID & "'", txtShortName.Text)
            curRemDesc_.Remdesc = Trim(txtShortName.Text)
            colRemDesc_.Add curRemDesc_, "'" & curRemDesc_.ID & "'"
        Else
            DBConnExc "Update tbRemarkList Set RemDesc='" & Trim(txtShortName.Text) & "' WHERE RemarkID='" & curRemDesc_.ID & "'"
            Set tmpList = lvList.ListItems("'" & curRemDesc_.ID & "'")
            curRemDesc_.Remdesc = Trim(txtShortName.Text)
        End If
        tmpList.Text = Trim(txtShortName.Text)
        Set tmpList = Nothing
        Set tmpCon = Nothing
    End If
    CreateNewRemDesc
    ClearScreen
    txtShortName.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim searchM As New BWGSearchManager
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    With frmPicture
        CmdSave.Picture = .picSave.Picture
        cmdNew.Picture = .picNew.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    Set colRemDesc_ = searchM.RemarkListSearch
    ShowRemarkList
    CreateNewRemDesc
End Sub

Private Sub ShowRemarkList()
Dim x%
Dim tmpList As ListItem
Dim tmpItem As BWGRemarkList
    If colRemDesc_ Is Nothing Then Exit Sub
    lvList.ListItems.Clear
    For x = 1 To colRemDesc_.Count
        Set tmpItem = colRemDesc_(x)
        Set tmpList = lvList.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.Remdesc)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
End Sub

Private Sub lvList_DblClick()
'    If Not callerForm_ Is Nothing And Not lvList.SelectedItem Is Nothing Then
'        If callerForm_.Visible Then
'            callerForm_.SetDataRemark colRemDesc_(lvList.SelectedItem.Key)
'            Unload Me
'        End If
'    End If
End Sub

Private Sub lvList_ItemClick(ByVal item As MSComctlLib.ListItem)
Dim tmpKey As String
    If lvList.SelectedItem Is Nothing Then Exit Sub
    Set curRemDesc_ = Nothing
    tmpKey = Trim(lvList.SelectedItem.Key)
    If isExist(colRemDesc_, tmpKey) Then
        Set curRemDesc_ = colRemDesc_(tmpKey)
        txtShortName.Text = curRemDesc_.Remdesc
        curRole_ = UpdateRole
    End If
End Sub

Private Sub CreateNewRemDesc()
    Set curRemDesc_ = New BWGRemarkList
    curRemDesc_.ID = ""
    curRole_ = AddNewRole
End Sub

Private Sub ClearScreen()
    txtShortName.Text = ""
End Sub

Private Sub lvList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpKey As String
Dim tmpStr As String
    'tmpStr = CurrentUser.VerifyAccessRole(Me.Name, DeleteRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    If lvList.SelectedItem Is Nothing Then Exit Sub
    If curRemDesc_.ID = "" Then
        MsgBox "กรุณาเลือกรายการที่ต้องการลบ", vbExclamation
        Exit Sub
    End If
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        tmpKey = "'" & curRemDesc_.ID & "'"
        If MsgBox("ต้องการลบข้อมูลรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            tmpStr = curRemDesc_.Delete
            If Trim(tmpStr) = "" Then
                If isExist(lvList.ListItems, tmpKey) Then lvList.ListItems.Remove tmpKey
                If isExist(colRemDesc_, tmpKey) Then colRemDesc_.Remove tmpKey
                CreateNewRemDesc
                ClearScreen
                txtShortName.SetFocus
            Else
                MsgBox tmpStr, vbExclamation
            End If
        End If
    End If
End Sub

Private Sub txtshortname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then CmdSave_Click
End Sub
