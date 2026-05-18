VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmLABTestType 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F052"
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
   Begin VB.TextBox txtPriority 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5790
      MaxLength       =   80
      TabIndex        =   1
      Top             =   870
      Width           =   675
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5190
      Width           =   1035
   End
   Begin VB.TextBox txtShortName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1410
      MaxLength       =   80
      TabIndex        =   0
      Top             =   870
      Width           =   2685
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5190
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5775
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5190
      Width           =   1035
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3825
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   6747
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
         Text            =   "ชื่อหน่วย"
         Object.Width           =   8466
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "ลำดับ"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ลำดับในการทดสอบ"
      Height          =   195
      Index           =   0
      Left            =   4380
      TabIndex        =   8
      Top             =   930
      Width           =   1350
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
      Caption         =   "ประเภทสำหรับห้อง LAB"
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
      TabIndex        =   7
      Top             =   0
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ชื่อประเภททดสอบ"
      Height          =   195
      Index           =   8
      Left            =   90
      TabIndex        =   6
      Top             =   930
      Width           =   1245
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frLABTestType.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6870
   End
End
Attribute VB_Name = "frmLABTestType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colTestType_ As Collection
Dim curTestType_ As BWGLABTestType
Dim curRole_ As AWSRoleType
Dim callerForm_ As Form

Private Function isChanged() As Boolean
Dim tmpVal As Integer
    If IsNumeric(txtPriority.Text) = False Then
        tmpVal = 0
    Else
        tmpVal = CInt(txtPriority.Text)
    End If
    If Trim(txtShortName.Text) <> Trim(curTestType_.TestTypeName) Or tmpVal <> curTestType_.TestPriority Then
        isChanged = True
    Else
        isChanged = False
    End If
End Function

Public Sub SetCallerForm(frm As Form, Optional setYear As Integer = 0)
    Set callerForm_ = frm
End Sub

Private Sub cmdClose_Click()
Dim aa
    If isChanged Then
        aa = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนปิดหรือไม่", vbQuestion + vbYesNoCancel)
        If aa = vbYes Then
            cmdSave_Click
            If isChanged Then
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
    If isChanged Then
        aa = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่", vbQuestion + vbYesNoCancel)
        If aa = vbYes Then
            cmdSave_Click
            If isChanged Then
                Exit Sub
            End If
        ElseIf aa = vbCancel Then
            Exit Sub
        End If
    End If
    CreateNewLABTestType
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

    If curTestType_ Is Nothing Then Exit Sub
    If Trim(txtShortName.Text) = vbNullString Then
        MsgBox "กรุณาใส่ชื่อประเภทการทดสอบก่อนเก็บข้อมูล", vbExclamation
        txtShortName.SetFocus
        Exit Sub
    End If
    If Trim(txtPriority.Text) = "" Then
        MsgBox "กรุณาใส่ข้อมูลลำดับในการทดสอบเท่านั้น", vbExclamation
        txtPriority.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtPriority.Text) = False Then
        MsgBox "กรุณาใส่ข้อมูลลำดับในการทดสอบเป็นตัวเลขเท่านั้น", vbExclamation
        txtPriority.SetFocus
        Exit Sub
    End If

    If isChanged Then
        If Trim(curTestType_.ID) = "" Then
            curTestType_.ID = GetGUID
            DBConnExc "Insert into tbLABTestType (TestTypeID,TestTypeName,TestPriority) Values ('" & Trim(curTestType_.ID) & "','" & (txtShortName.Text) & "'," & Trim(txtPriority.Text) & ")"
            Set tmpList = lvList.ListItems.Add(, "'" & curTestType_.ID & "'", txtShortName.Text)
            curTestType_.TestTypeName = Trim(txtShortName.Text)
            curTestType_.TestPriority = Trim(txtPriority.Text)
            colTestType_.Add curTestType_, "'" & curTestType_.ID & "'"
        Else
            DBConnExc "Update tbLABTestType Set TestTypeName='" & Trim(txtShortName.Text) & "',TestPriority=" & txtPriority.Text & " WHERE TestTypeID='" & curTestType_.ID & "'"
            Set tmpList = lvList.ListItems("'" & curTestType_.ID & "'")
            curTestType_.TestTypeName = Trim(txtShortName.Text)
            curTestType_.TestPriority = Trim(txtPriority.Text)
        End If
        tmpList.Text = Trim(txtShortName.Text)
        tmpList.SubItems(1) = Trim(txtPriority.Text)
        Set tmpList = Nothing
        Set tmpCon = Nothing
    End If
    CreateNewLABTestType
    ClearScreen
    txtShortName.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
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
    
Dim searchM As New BWGSearchManager
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdNew.Picture = .picNew.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    Set colTestType_ = searchM.TestTypeLABSearch
    ShowUnitList
    CreateNewLABTestType
End Sub

Private Sub ShowUnitList()
Dim x%
Dim tmpList As ListItem
Dim tmpItem As BWGLABTestType
    If colTestType_ Is Nothing Then Exit Sub
    lvList.ListItems.Clear
    For x = 1 To colTestType_.Count
        Set tmpItem = colTestType_(x)
        Set tmpList = lvList.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.TestTypeName)
        tmpList.SubItems(1) = tmpItem.TestPriority
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
End Sub

Private Sub lvList_DblClick()
On Error Resume Next
    If Not callerForm_ Is Nothing And Not lvList.SelectedItem Is Nothing Then
        callerForm_.SetDataUnit colTestType_(lvList.SelectedItem.Key)
        Unload Me
    End If
End Sub

Private Sub lvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim tmpKey As String
    If lvList.SelectedItem Is Nothing Then Exit Sub
    Set curTestType_ = Nothing
    tmpKey = Trim(lvList.SelectedItem.Key)
    If isExist(colTestType_, tmpKey) Then
        Set curTestType_ = colTestType_(tmpKey)
        txtShortName.Text = curTestType_.TestTypeName
        txtPriority.Text = curTestType_.TestPriority
        curRole_ = UpdateRole
    End If
End Sub

Private Sub CreateNewLABTestType()
    Set curTestType_ = New BWGLABTestType
    curTestType_.ID = ""
    curRole_ = AddNewRole
End Sub

Private Sub ClearScreen()
    txtShortName.Text = ""
    txtPriority.Text = ""
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
    If curTestType_.ID = "" Then
        MsgBox "กรุณาเลือกรายการที่ต้องการลบ", vbExclamation
        Exit Sub
    End If
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        tmpKey = "'" & curTestType_.ID & "'"
        If MsgBox("ต้องการลบข้อมูลรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            tmpStr = curTestType_.Delete
            If Trim(tmpStr) = "" Then
                If isExist(lvList.ListItems, tmpKey) Then lvList.ListItems.Remove tmpKey
                If isExist(colTestType_, tmpKey) Then colTestType_.Remove tmpKey
                CreateNewLABTestType
                ClearScreen
                txtShortName.SetFocus
            Else
                MsgBox tmpStr, vbExclamation
            End If
        End If
    End If
End Sub
