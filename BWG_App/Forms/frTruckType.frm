VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmTruckType 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F134"
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
   Begin VB.CheckBox chkFactoryCar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "รถโรงงาน"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   5580
      TabIndex        =   2
      Top             =   870
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chkTrailer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "เป็นรถพ่วงได้"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4020
      TabIndex        =   1
      Top             =   870
      Width           =   1395
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
      Left            =   1200
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ประเภทรถ"
         Object.Width           =   8466
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เป็นรถพ่วงได้"
         Object.Width           =   2540
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
      Caption         =   "ประเภทของรถบรรทุก"
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
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ชื่อประเภทรถ"
      Height          =   195
      Index           =   8
      Left            =   210
      TabIndex        =   7
      Top             =   930
      Width           =   945
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frTruckType.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6870
   End
End
Attribute VB_Name = "frmTruckType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim coltruckType_ As Collection
Dim curTruckType_ As BWGTruckType
Dim curRole_ As AWSRoleType
Dim callerForm_ As Form

Public Sub SetCallerForm(frm As Form, Optional setYear As Integer = 0)
    Set callerForm_ = frm
End Sub

Private Function isChanged() As Boolean
Dim tmpVal As Boolean
Dim tmpVal2 As Boolean
    tmpVal = (chkTrailer.Value = 1)
    tmpVal2 = (chkFactoryCar.Value = 1)
    If Trim(txtShortName.Text) <> Trim(curTruckType_.TypeDesc) Or tmpVal <> curTruckType_.CanBeTrailer Or tmpVal2 <> curTruckType_.isFactoryCar Then
        isChanged = True
    Else
        isChanged = False
    End If
End Function

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
    txtShortName.SetFocus
End Sub

Private Sub cmdSave_Click()
Dim sqlStr As String
Dim tmpList As ListItem, tmpVal2 As Boolean
Dim tmpCon As ADODB.Connection
Dim tmpStr As String, tmpVal As Boolean
    'tmpStr = CurrentUser.VerifyAccessRole(Me.Name, curRole_)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If

    If curTruckType_ Is Nothing Then Exit Sub
    If Trim(txtShortName.Text) = vbNullString Then
        MsgBox "กรุณาใส่ชื่อประเภทรถก่อนเก็บข้อมูล", vbExclamation
        txtShortName.SetFocus
        Exit Sub
    End If

    If isChanged Then
        tmpVal = (chkTrailer.Value = 1)
        tmpVal2 = (chkFactoryCar.Value = 1)
        If Trim(curTruckType_.ID) = "" Then
            curTruckType_.ID = GetGUID
            DBConnExc "Insert into tbTruckType (TruckTypeID,TruckTypeDesc,CanBeTrailer,isFactoryCar) Values ('" & Trim(curTruckType_.ID) & "','" & (txtShortName.Text) & "','" & TF_To_YN(tmpVal) & "','" & TF_To_YN(tmpVal2) & "')"
            Set tmpList = lvList.ListItems.Add(, "'" & curTruckType_.ID & "'", txtShortName.Text)
            curTruckType_.TypeDesc = Trim(txtShortName.Text)
            curTruckType_.CanBeTrailer = tmpVal
            curTruckType_.isFactoryCar = tmpVal2
            coltruckType_.Add curTruckType_, "'" & curTruckType_.ID & "'"
        Else
            DBConnExc "Update tbTruckType Set TruckTypeDesc='" & Trim(txtShortName.Text) & "',CanBeTrailer='" & TF_To_YN(tmpVal) & "',isFactoryCar='" & TF_To_YN(tmpVal2) & "'  WHERE TruckTypeID='" & curTruckType_.ID & "'"
            Set tmpList = lvList.ListItems("'" & curTruckType_.ID & "'")
            curTruckType_.TypeDesc = Trim(txtShortName.Text)
            curTruckType_.CanBeTrailer = tmpVal
            curTruckType_.isFactoryCar = tmpVal2
        End If
        tmpList.Text = Trim(txtShortName.Text)
        If tmpVal Then
            tmpList.SubItems(1) = "Yes"
        Else
            tmpList.SubItems(1) = "No"
        End If
        Set tmpList = Nothing
        Set tmpCon = Nothing
    End If
    CreateNewTruckType
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
        cmdSave.Picture = .picSave.Picture
        cmdNew.Picture = .picNew.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    Set coltruckType_ = searchM.TruckTypeSearch
    ShowTruckTypeList
    CreateNewTruckType
End Sub

Private Sub ShowTruckTypeList()
Dim x%
Dim tmpList As ListItem
Dim tmpItem As BWGTruckType
    If coltruckType_ Is Nothing Then Exit Sub
    lvList.ListItems.Clear
    For x = 1 To coltruckType_.Count
        Set tmpItem = coltruckType_(x)
        Set tmpList = lvList.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.TypeDesc)
        If tmpItem.CanBeTrailer Then
            tmpList.SubItems(1) = "Yes"
        Else
            tmpList.SubItems(1) = "No"
        End If
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
End Sub

Private Sub lvList_DblClick()
On Error Resume Next
    If Not callerForm_ Is Nothing And Not lvList.SelectedItem Is Nothing Then
        callerForm_.SetDataTruckType coltruckType_(lvList.SelectedItem.Key)
        Unload Me
    End If
End Sub

Private Sub lvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim tmpKey As String
    If lvList.SelectedItem Is Nothing Then Exit Sub
    Set curTruckType_ = Nothing
    tmpKey = Trim(lvList.SelectedItem.Key)
    If isExist(coltruckType_, tmpKey) Then
        Set curTruckType_ = coltruckType_(tmpKey)
        txtShortName.Text = curTruckType_.TypeDesc
        If curTruckType_.CanBeTrailer Then
            chkTrailer.Value = 1
        Else
            chkTrailer.Value = 0
        End If
        If curTruckType_.isFactoryCar Then
            chkFactoryCar.Value = 1
        Else
            chkFactoryCar.Value = 0
        End If
        curRole_ = UpdateRole
    End If
End Sub

Private Sub CreateNewTruckType()
    Set curTruckType_ = New BWGTruckType
    curTruckType_.ID = ""
    curRole_ = AddNewRole
End Sub

Private Sub ClearScreen()
    txtShortName.Text = ""
    chkTrailer.Value = 0
    chkFactoryCar.Value = 0
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
    If curTruckType_.ID = "" Then
        MsgBox "กรุณาเลือกรายการที่ต้องการลบ", vbExclamation
        Exit Sub
    End If
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        tmpKey = "'" & curTruckType_.ID & "'"
        If MsgBox("ต้องการลบข้อมูลรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            tmpStr = curTruckType_.Delete
            If Trim(tmpStr) = "" Then
                If isExist(lvList.ListItems, tmpKey) Then lvList.ListItems.Remove tmpKey
                If isExist(coltruckType_, tmpKey) Then coltruckType_.Remove tmpKey
                CreateNewTruckType
                ClearScreen
                txtShortName.SetFocus
            Else
                MsgBox tmpStr, vbExclamation
            End If
        End If
    End If
End Sub

Private Sub txtshortname_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub
