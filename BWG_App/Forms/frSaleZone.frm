VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmSaleZone 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F117"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSaleStaff 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3180
      MaxLength       =   80
      TabIndex        =   9
      Top             =   1260
      Width           =   3615
   End
   Begin VB.TextBox txtZoneDesc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3180
      MaxLength       =   80
      TabIndex        =   1
      Top             =   900
      Width           =   3615
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5220
      Width           =   1035
   End
   Begin VB.TextBox txtShortName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   900
      MaxLength       =   80
      TabIndex        =   0
      Top             =   900
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5220
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5220
      Width           =   1035
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3525
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1620
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   6218
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ชื่อ Zone"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "คำอธิบาย"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "หัวหน้าส่วนการตลาด"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "สถานะ"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "เซลล์ที่รับผิดชอบ"
      Height          =   195
      Index           =   2
      Left            =   1920
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "คำอธิบาย"
      Height          =   195
      Index           =   0
      Left            =   2460
      TabIndex        =   8
      Top             =   960
      Width           =   675
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
      Caption         =   "กลุ่ม Zone ของเจ้าหน้าที่การตลาด"
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
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ชื่อ Zone"
      Height          =   195
      Index           =   8
      Left            =   180
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frSaleZone.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6870
   End
End
Attribute VB_Name = "frmSaleZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colZone_ As Collection
Dim curZone_ As BWGSaleZone
Dim curRole_ As AWSRoleType
Dim callerForm_ As Form
Dim curStaff_ As BWGCompanyStaff

Public Sub SetCallerForm(frm As Form, Optional setYear As Integer = 0)
    Set callerForm_ = frm
End Sub

Private Sub cmdClose_Click()
Dim aa
    If Trim(txtShortName.Text) <> Trim(curZone_.ZoneName) Or Trim(txtZoneDesc.Text) <> Trim(curZone_.Description) Then
        aa = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนปิดหรือไม่", vbQuestion + vbYesNoCancel)
        If aa = vbYes Then
            cmdSave_Click
            If Trim(txtShortName.Text) <> Trim(curZone_.ZoneName) Or Trim(txtZoneDesc.Text) <> Trim(curZone_.Description) Then
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
    If Trim(txtShortName.Text) <> Trim(curZone_.ZoneName) Or Trim(txtZoneDesc.Text) <> Trim(curZone_.Description) Then
        aa = MsgBox("ข้อมูลมีการเปลี่ยนแปลง ต้องการจัดเก็บข้อมูลก่อนสร้างใหม่หรือไม่", vbQuestion + vbYesNoCancel)
        If aa = vbYes Then
            cmdSave_Click
            If Trim(txtShortName.Text) <> Trim(curZone_.ZoneName) Or Trim(txtZoneDesc.Text) <> Trim(curZone_.Description) Then
                Exit Sub
            End If
        ElseIf aa = vbCancel Then
            Exit Sub
        End If
    End If
    CreateNewZone
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

    If curZone_ Is Nothing Then Exit Sub
    If Trim(txtShortName.Text) = vbNullString Then
        MsgBox "กรุณาใส่ชื่อ Zone ก่อนเก็บข้อมูล", vbExclamation
        txtShortName.SetFocus
        Exit Sub
    End If

    If Trim(txtShortName.Text) <> Trim(curZone_.ZoneName) Or Trim(txtZoneDesc.Text) <> Trim(curZone_.Description) Or Trim(txtSaleStaff.Text) <> Trim(curZone_.SupName) Then
        If Trim(curZone_.ID) = "" Then
            curZone_.ID = GetGUID
            DBConnExc "Insert into tbSaleZone (SaleZoneID,ZoneName,ZoneDesc, SupID, isActive) Values ('" & Trim(curZone_.ID) & "','" & (txtShortName.Text) & "','" & Trim(txtZoneDesc.Text) & "', '" & curStaff_.ID & "', '" & IIf(lvList.SelectedItem.Checked = True, "Y", "N") & "')"
            Set tmpList = lvList.ListItems.Add(, "'" & curZone_.ID & "'", txtShortName.Text)
            curZone_.ZoneName = Trim(txtShortName.Text)
            curZone_.Description = Trim(txtZoneDesc.Text)
            colZone_.Add curZone_, "'" & curZone_.ID & "'"
        Else
            DBConnExc "Update tbSaleZone Set ZoneName='" & Trim(txtShortName.Text) & "',ZoneDesc='" & txtZoneDesc.Text & "', SupID = '" & curStaff_.ID & "', isActive = '" & IIf(lvList.SelectedItem.Checked = True, "Y", "N") & "' WHERE SaleZoneID='" & curZone_.ID & "'"
            Set tmpList = lvList.ListItems("'" & curZone_.ID & "'")
            curZone_.ZoneName = Trim(txtShortName.Text)
            curZone_.Description = Trim(txtZoneDesc.Text)
        End If
        tmpList.Text = Trim(txtShortName.Text)
        tmpList.SubItems(1) = txtZoneDesc.Text
        tmpList.SubItems(2) = txtSaleStaff.Text
        tmpList.SubItems(3) = IIf(lvList.SelectedItem.Checked = True, "Y", "N")
        Set tmpList = Nothing
        Set tmpCon = Nothing
    End If
    CreateNewZone
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
    Set colZone_ = searchM.SaleZoneSearch
    ListZone
    CreateNewZone
End Sub

Private Sub ListZone()
Dim x%
Dim tmpList As ListItem
Dim tmpItem As BWGSaleZone
    If colZone_ Is Nothing Then Exit Sub
    lvList.ListItems.Clear
    For x = 1 To colZone_.Count
        Set tmpItem = colZone_(x)
        Set tmpList = lvList.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.ZoneName)
        tmpList.SubItems(1) = tmpItem.Description
        tmpList.SubItems(2) = tmpItem.SupName
        tmpList.SubItems(3) = tmpItem.isActive
        If UCase(tmpItem.isActive) = "Y" Then
            tmpList.Checked = True
        Else
            tmpList.Checked = False
        End If
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
End Sub

Private Sub lvList_DblClick()
    If Not callerForm_ Is Nothing And Not lvList.SelectedItem Is Nothing Then
        callerForm_.SetDataUnit colZone_(lvList.SelectedItem.Key)
        Unload Me
    End If
End Sub

Private Sub lvList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim tSql As String
    If Item.Checked = False Then
        tSql = "Exec sp_UpdateZoneStatus " & Item.Key & ", 'N'"
        DBConnExc tSql
        Item.SubItems(3) = "N"
    ElseIf Item.Checked = True Then
        tSql = "Exec sp_UpdateZoneStatus " & Item.Key & ", 'Y'"
        DBConnExc tSql
        Item.SubItems(3) = "Y"
    End If
End Sub

Private Sub lvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim tmpKey As String
    If lvList.SelectedItem Is Nothing Then Exit Sub
    Set curZone_ = Nothing
    tmpKey = Trim(lvList.SelectedItem.Key)
    If isExist(colZone_, tmpKey) Then
        Set curZone_ = colZone_(tmpKey)
        txtShortName.Text = curZone_.ZoneName
        txtZoneDesc.Text = curZone_.Description
        txtSaleStaff.Text = curZone_.SupName
        curRole_ = UpdateRole
    End If
End Sub

Private Sub CreateNewZone()
    Set curZone_ = New BWGSaleZone
    curZone_.ID = ""
    curRole_ = AddNewRole
End Sub

Private Sub ClearScreen()
    txtShortName.Text = ""
    txtZoneDesc.Text = ""
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
    If curZone_.ID = "" Then
        MsgBox "กรุณาเลือกรายการที่ต้องการลบ", vbExclamation
        Exit Sub
    End If
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        tmpKey = "'" & curZone_.ID & "'"
        If MsgBox("ต้องการลบข้อมูลรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            tmpStr = curZone_.Delete
            If Trim(tmpStr) = "" Then
                If isExist(lvList.ListItems, tmpKey) Then lvList.ListItems.Remove tmpKey
                If isExist(colZone_, tmpKey) Then colZone_.Remove tmpKey
                CreateNewZone
                ClearScreen
                txtShortName.SetFocus
            Else
                MsgBox tmpStr, vbExclamation
            End If
        End If
    End If
End Sub

Private Sub txtSaleStaff_Validate(Cancel As Boolean)
    If curStaff_ Is Nothing Then Set curStaff_ = New BWGCompanyStaff
    If Trim(txtSaleStaff.Text) = "" And Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSaleStaff.Text) Then
        Set curStaff_ = Nothing
        Set curStaff_ = New BWGCompanyStaff
    Else
        If Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName) <> Trim(txtSaleStaff.Text) Then
            Set curStaff_ = Nothing
            Set curStaff_ = New BWGCompanyStaff
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtSaleStaff.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If curStaff_.ID = "" Then
                txtSaleStaff.Text = ""
            End If
        End If
    End If
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff)
    Set curStaff_ = selStaff
    With curStaff_
        txtSaleStaff.Text = .StaffFName & " " & .StaffLName
    End With
End Sub

