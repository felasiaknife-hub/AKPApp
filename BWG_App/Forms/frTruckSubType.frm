VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmTruckSubType 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F133"
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
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "คิดราคาตามน้ำหนักบรรทุก"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   4530
      TabIndex        =   9
      Top             =   1200
      Width           =   2175
   End
   Begin VB.ComboBox cboTruckType 
      Height          =   315
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   810
      Width           =   2655
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
      Left            =   1740
      MaxLength       =   80
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
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
      Height          =   3465
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1650
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   6112
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
         Text            =   "กลุ่มหลัก"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "กลุ่มย่อย"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "การคิดราคา"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "กลุ่มหลัก"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   1050
      TabIndex        =   8
      Top             =   870
      Width           =   600
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
      Caption         =   "ประเภทกลุ่มย่อยของรถบรรทุก"
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
      Width           =   2625
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ชื่อประเภทกลุ่มย่อยรถ"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   6
      Top             =   1260
      Width           =   1545
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frTruckSubType.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6870
   End
End
Attribute VB_Name = "frmTruckSubType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim coltruckType_ As Collection
Dim colSubTruckType_ As Collection
Dim curSubTruckType_ As BWGTruckSubType
Dim curRole_ As AWSRoleType
Dim callerForm_ As Form

Public Sub SetCallerForm(frm As Form, Optional setYear As Integer = 0)
    Set callerForm_ = frm
End Sub

Private Sub cboTruckType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cboTruckType_Validate(Cancel As Boolean)
    If cboTruckType.ListIndex > 0 Then
        curSubTruckType_.TruckTypeID = coltruckType_(cboTruckType.ListIndex).ID
        curSubTruckType_.TypeDesc = coltruckType_(cboTruckType.ListIndex).TypeDesc
    Else
        curSubTruckType_.TruckTypeID = ""
        curSubTruckType_.TypeDesc = ""
    End If
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
    CreateNewTruckType
    ClearScreen
    cboTruckType.SetFocus
End Sub

Private Sub cmdSave_Click()
Dim sqlStr As String
Dim tmpList As ListItem
Dim tmpCon As ADODB.Connection
Dim tmpStr As String, isChargeByWeight As Boolean
    'tmpStr = CurrentUser.VerifyAccessRole(Me.Name, curRole_)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If

    If curSubTruckType_ Is Nothing Then Exit Sub
    If Trim(txtShortName.Text) = vbNullString Then
        MsgBox "กรุณาใส่ชื่อประเภทรถก่อนเก็บข้อมูล", vbExclamation
        txtShortName.SetFocus
        Exit Sub
    End If
    If cboTruckType.ListIndex = 0 Then
        MsgBox "กรุณาระบุประเภทกลุ่มหลักของรถ", vbExclamation
        cboTruckType.SetFocus
        Exit Sub
    End If
    If Check1.Value = 1 Then
        isChargeByWeight = True
    Else
        isChargeByWeight = False
    End If
    If isChanged Then
        If Trim(curSubTruckType_.ID) = "" Then
            curSubTruckType_.ID = GetGUID
            DBConnExc "Insert into tbTruckSubType (SubTypeID, TruckTypeID, SubTypeDesc,isChargeByWeight) Values ('" & Trim(curSubTruckType_.ID) & "','" & Trim(curSubTruckType_.TruckTypeID) & "','" & (txtShortName.Text) & "','" & TF_To_YN(isChargeByWeight) & "')"
            Set tmpList = lvList.ListItems.Add(, "'" & curSubTruckType_.ID & "'", txtShortName.Text)
            curSubTruckType_.SubTypeDesc = Trim(txtShortName.Text)
            colSubTruckType_.Add curSubTruckType_, "'" & curSubTruckType_.ID & "'"
        Else
           DBConnExc "Update tbTruckSubType Set TruckTypeID='" & Trim(curSubTruckType_.TruckTypeID) & "', " & " SubTypeDesc='" & Trim(txtShortName.Text) & "',isChargeByWeight='" & TF_To_YN(isChargeByWeight) & "' WHERE SubTypeID='" & curSubTruckType_.ID & "'"
            Set tmpList = lvList.ListItems("'" & curSubTruckType_.ID & "'")
            curSubTruckType_.SubTypeDesc = Trim(txtShortName.Text)
        End If
        curSubTruckType_.isChargeByWeight = isChargeByWeight
        tmpList.Text = Trim(cboTruckType.Text)
        tmpList.SubItems(1) = Trim(txtShortName.Text)
        If curSubTruckType_.isChargeByWeight Then
            tmpList.SubItems(2) = "คิดตามน้ำหนัก"
        Else
            tmpList.SubItems(2) = "คิดตามระยะทาง"
        End If
        Set tmpList = Nothing
        Set tmpCon = Nothing
    End If
    CreateNewTruckType
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
Dim searchM As New BWGSearchManager
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    With frmPicture
        cmdSave.Picture = .picSave.Picture
        cmdNew.Picture = .picNew.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    PopulateTruckType
    Set colSubTruckType_ = searchM.TruckSubTypeSearch
    ShowTruckTypeList
    CreateNewTruckType
End Sub

Private Sub ShowTruckTypeList()
Dim x%
Dim tmpList As ListItem
Dim tmpItem As BWGTruckSubType
    If colSubTruckType_ Is Nothing Then Exit Sub
    lvList.ListItems.Clear
    For x = 1 To colSubTruckType_.Count
        Set tmpItem = colSubTruckType_(x)
        Set tmpList = lvList.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.TypeDesc)
        tmpList.SubItems(1) = tmpItem.SubTypeDesc
        If tmpItem.isChargeByWeight Then
            tmpList.SubItems(2) = "คิดตามน้ำหนัก"
        Else
            tmpList.SubItems(2) = "คิดตามระยะทาง"
        End If
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
End Sub

Private Sub lvList_DblClick()
On Error Resume Next
    If Not callerForm_ Is Nothing And Not lvList.SelectedItem Is Nothing Then
        callerForm_.SetDataTruckType colSubTruckType_(lvList.SelectedItem.Key)
        Unload Me
    End If
End Sub

Private Sub lvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim tmpKey As String
    If lvList.SelectedItem Is Nothing Then Exit Sub
    Set curSubTruckType_ = Nothing
    tmpKey = Trim(lvList.SelectedItem.Key)
    If isExist(colSubTruckType_, tmpKey) Then
        Set curSubTruckType_ = colSubTruckType_(tmpKey)
        txtShortName.Text = curSubTruckType_.SubTypeDesc
        SetTextToCombo curSubTruckType_.TypeDesc, cboTruckType
        If curSubTruckType_.isChargeByWeight Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
        curRole_ = UpdateRole
    End If
End Sub

Private Sub CreateNewTruckType()
    Set curSubTruckType_ = New BWGTruckSubType
    curSubTruckType_.ID = ""
    curSubTruckType_.isChargeByWeight = False
    curSubTruckType_.isChanged = False
    curRole_ = AddNewRole
End Sub

Private Sub ClearScreen()
    txtShortName.Text = ""
    cboTruckType.ListIndex = 0
    Check1.Value = 0
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
    If curSubTruckType_.ID = "" Then
        MsgBox "กรุณาเลือกรายการที่ต้องการลบ", vbExclamation
        Exit Sub
    End If
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        tmpKey = "'" & curSubTruckType_.ID & "'"
        If MsgBox("ต้องการลบข้อมูลรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            tmpStr = curSubTruckType_.Delete
            If Trim(tmpStr) = "" Then
                If isExist(lvList.ListItems, tmpKey) Then lvList.ListItems.Remove tmpKey
                If isExist(colSubTruckType_, tmpKey) Then colSubTruckType_.Remove tmpKey
                CreateNewTruckType
                ClearScreen
                txtShortName.SetFocus
            Else
                MsgBox tmpStr, vbExclamation
            End If
        End If
    End If
End Sub

Private Sub PopulateTruckType()
Dim tmpS As New BWGSearchManager
Dim x%
        
    Set coltruckType_ = tmpS.TruckTypeSearch
    cboTruckType.Clear
    cboTruckType.AddItem "== กรุณาระบุ ประเภทรถ =="
    For x = 1 To coltruckType_.Count
        cboTruckType.AddItem coltruckType_(x).TypeDesc
    Next
    cboTruckType.ListIndex = 0
End Sub

Private Function isChanged() As Boolean
Dim isChargeByWeight As Boolean
    isChargeByWeight = Check1.Value = 1
    If Trim(txtShortName.Text) <> Trim(curSubTruckType_.SubTypeDesc) Or isChargeByWeight <> curSubTruckType_.isChargeByWeight Then
        isChanged = True
    Else
        isChanged = False
    End If
End Function
