VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmLABParam 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F046"
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
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5100
      MaxLength       =   80
      TabIndex        =   1
      Top             =   870
      Width           =   855
   End
   Begin VB.ComboBox cboUnit 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5490
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5190
      Width           =   1035
   End
   Begin VB.TextBox txtShortName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   480
      MaxLength       =   80
      TabIndex        =   0
      Top             =   870
      Width           =   4035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5190
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5775
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5190
      Width           =   1035
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3825
      Left            =   0
      TabIndex        =   3
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
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ชื่อ"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "หน่วย"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "รหัส"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "รหัส"
      Height          =   195
      Index           =   2
      Left            =   4740
      TabIndex        =   10
      Top             =   930
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "หน่วย"
      Height          =   195
      Index           =   0
      Left            =   1380
      TabIndex        =   9
      Top             =   5550
      Visible         =   0   'False
      Width           =   420
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
      Caption         =   "พารามิเตอร์สำหรับ LAB"
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
      TabIndex        =   8
      Top             =   0
      Width           =   2070
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ชื่อ"
      Height          =   195
      Index           =   8
      Left            =   210
      TabIndex        =   7
      Top             =   930
      Width           =   195
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frLABParam.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6870
   End
End
Attribute VB_Name = "frmLABParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colLABParam_ As Collection
Dim curLABParam_ As BWGLABParameterList
Dim curRole_ As AWSRoleType
Dim callerForm_ As Form

Public Sub SetCallerForm(frm As Form, Optional setYear As Integer = 0)
    Set callerForm_ = frm
End Sub

Private Sub cboUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        cboUnit.ListIndex = -1
    End If
End Sub

Private Sub cboUnit_Validate(Cancel As Boolean)
    If cboUnit.ListIndex > 0 Then
        curLABParam_.LabUnitID = colUnit_(cboUnit.ListIndex + 1).ID
    Else
        curLABParam_.LabUnitID = ""
    End If
End Sub

Private Sub cmdClose_Click()
Dim aa
    If Trim(txtShortName.Text) <> Trim(curLABParam_.ParameterName) Then
        aa = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนปิดหรือไม่", vbQuestion + vbYesNoCancel)
        If aa = vbYes Then
            cmdSave_Click
        If Trim(txtShortName.Text) <> Trim(curLABParam_.ParameterName) Then
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
    If Trim(txtShortName.Text) <> Trim(curLABParam_.ParameterName) Then
        aa = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่", vbQuestion + vbYesNoCancel)
        If aa = vbYes Then
            cmdSave_Click
            If Trim(txtShortName.Text) <> Trim(curLABParam_.ParameterName) Then
                Exit Sub
            End If
        ElseIf aa = vbCancel Then
            Exit Sub
        End If
    End If
    CreateNewLABParam
    ClearScreen
    txtShortName.SetFocus
End Sub

Private Sub cmdSave_Click()
Dim sqlStr As String
Dim tmpList As ListItem
Dim tmpCon As ADODB.Connection
Dim tmpStr As String
    'tmpStr = CurrentUser.VerifyAccessRole(Me.Name, curRole_)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If

    If curLABParam_ Is Nothing Then Exit Sub

    If isChanged Then
        If Trim(curLABParam_.ID) = "" Then
            curLABParam_.ID = GetGUID
            DBConnExc "Insert into tbLABParameterList (ParameterID, ParameterName, LabUnitID,ParamCode) Values ('" & Trim(curLABParam_.ID) & "','" & Trim(txtShortName.Text) & "','" & Trim(curLABParam_.LabUnitID) & "','" & Trim(txtCode.Text) & "')"
            Set tmpList = lvList.ListItems.Add(, "'" & curLABParam_.ID & "'", txtShortName.Text)
            curLABParam_.ParameterName = Trim(txtShortName.Text)
            curLABParam_.ParameterCode = Trim(txtCode.Text)
            colLABParam_.Add curLABParam_, "'" & curLABParam_.ID & "'"
        Else
            DBConnExc "Update tbLABParameterList Set ParameterName='" & Trim(txtShortName.Text) & "',LabUnitID='" & curLABParam_.LabUnitID & "',ParamCode='" & Trim(txtCode.Text) & "' WHERE ParameterID='" & curLABParam_.ID & "'"
            Set tmpList = lvList.ListItems("'" & curLABParam_.ID & "'")
            curLABParam_.ParameterName = Trim(txtShortName.Text)
            curLABParam_.ParameterCode = Trim(txtCode.Text)
            curLABParam_.LabUnitID = curLABParam_.LabUnitID
        End If
        tmpList.Text = Trim(txtShortName.Text)
        tmpList.SubItems(1) = Trim(cboUnit.Text)
        tmpList.SubItems(2) = Trim(txtCode.Text)
        Set tmpList = Nothing
        Set tmpCon = Nothing
    End If
    CreateNewLABParam
    ClearScreen
    txtShortName.SetFocus
End Sub

Private Sub PopulateLABParam()
Dim x%
Dim tmpS As New BWGSearchManager
    Set colUnit_ = tmpS.UnitLABSearch
    cboUnit.Clear
    For x = 1 To colUnit_.Count
        cboUnit.AddItem colUnit_(x).UnitName
    Next
    Set tmpS = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim searchM As New BWGSearchManager
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdNew.Picture = .picNew.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    
    Set colLABParam_ = searchM.ParameterLABSearch
    
    Set curLABParam_ = Nothing
    Set curLABParam_ = New BWGLABParameterList

    PopulateLABParam
    ShowLABParamList
    CreateNewLABParam
End Sub

Private Sub ShowLABParamList()
Dim x%
Dim tmpList As ListItem
Dim tmpItem As BWGLABParameterList
    If colLABParam_ Is Nothing Then Exit Sub
    lvList.ListItems.Clear
    For x = 1 To colLABParam_.Count
        Set tmpItem = colLABParam_(x)
        Set tmpList = lvList.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.ParameterName)
        tmpList.SubItems(1) = Trim(tmpItem.LabUnitName)
        tmpList.SubItems(2) = Trim(tmpItem.ParameterCode)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
End Sub

Private Sub lvList_DblClick()
On Error Resume Next
    If Not callerForm_ Is Nothing And Not lvList.SelectedItem Is Nothing Then
        callerForm_.ShowLABParamList colLABParam_(lvList.SelectedItem.Key)
        Unload Me
    End If
End Sub

Private Sub lvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim tmpKey As String
    If lvList.SelectedItem Is Nothing Then Exit Sub
    Set curLABParam_ = Nothing
    tmpKey = Trim(lvList.SelectedItem.Key)
    If isExist(colLABParam_, tmpKey) Then
        Set curLABParam_ = colLABParam_(tmpKey)
        txtShortName.Text = curLABParam_.ParameterName
        txtCode.Text = curLABParam_.ParameterCode
'        cboUnit.Text = curLABParam_.LabUnitName
        curRole_ = UpdateRole
    End If
End Sub

Private Sub CreateNewLABParam()
    Set curLABParam_ = New BWGLABParameterList
    curLABParam_.ID = ""
    curRole_ = AddNewRole
End Sub

Private Sub ClearScreen()
    txtShortName.Text = ""
    cboUnit.ListIndex = -1
    txtCode.Text = ""
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
    If curLABParam_.ID = "" Then
        MsgBox "กรุณาเลือกรายการที่ต้องการลบ", vbExclamation
        Exit Sub
    End If
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        tmpKey = "'" & curLABParam_.ID & "'"
        If MsgBox("ต้องการลบข้อมูลรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            tmpStr = curLABParam_.Delete
            If Trim(tmpStr) = "" Then
                If isExist(lvList.ListItems, tmpKey) Then lvList.ListItems.Remove tmpKey
                If isExist(colLABParam_, tmpKey) Then colLABParam_.Remove tmpKey
                CreateNewLABParam
                ClearScreen
                txtShortName.SetFocus
            Else
                MsgBox tmpStr, vbExclamation
            End If
        End If
    End If
End Sub

Private Function isChanged() As Boolean
    If Trim(txtShortName.Text) <> Trim(curLABParam_.ParameterName) Or Trim(cboUnit.Text) <> Trim(curLABParam_.LabUnitName) Or Trim(txtCode.Text) <> Trim(curLABParam_.ParameterCode) Then
        isChanged = True
    Else
        isChanged = False
    End If
End Function
