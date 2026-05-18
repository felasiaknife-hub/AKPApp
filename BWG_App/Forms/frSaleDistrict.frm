VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmSaleDistrict 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "F114"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6840
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboZone 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   810
      Width           =   2895
   End
   Begin VB.TextBox txtSaleStaff 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3720
      MaxLength       =   80
      TabIndex        =   2
      Top             =   1200
      Width           =   3045
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "สร้างใหม่"
      Height          =   795
      Left            =   4710
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1035
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   900
      MaxLength       =   80
      TabIndex        =   1
      Top             =   1200
      Width           =   1425
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   5775
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5520
      Width           =   1035
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3825
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1650
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   6747
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ชื่อเขต"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เซลล์ที่รับผิดชอบ"
         Object.Width           =   7232
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "สถานะ"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "โซน"
      Height          =   195
      Index           =   2
      Left            =   540
      TabIndex        =   10
      Top             =   870
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "เซลล์ที่รับผิดชอบ"
      Height          =   195
      Index           =   0
      Left            =   2460
      TabIndex        =   9
      Top             =   1260
      Width           =   1170
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
      Caption         =   "เขตในโซน"
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
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ชื่อเขต"
      Height          =   195
      Index           =   8
      Left            =   360
      TabIndex        =   7
      Top             =   1260
      Width           =   465
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frSaleDistrict.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6870
   End
   Begin VB.Menu MNU 
      Caption         =   "เมนู"
      Begin VB.Menu MNU_All 
         Caption         =   "แสดงทั้งหมด"
      End
      Begin VB.Menu MNU_Active 
         Caption         =   "แสดงเฉพาะที่ Active"
      End
   End
End
Attribute VB_Name = "frmSaleDistrict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colDistrict_ As Collection
Dim curDistrict_ As BWGSaleDistrict
Dim colZone_ As Collection
Dim curStaff_ As BWGCompanyStaff
Dim curRole_ As AWSRoleType
Dim callerForm_ As Form
Dim tSql As String

Public Sub SetCallerForm(frm As Form, Optional setYear As Integer = 0)
    Set callerForm_ = frm
End Sub

Public Function isChanged() As Boolean
    If (Trim(cboZone.Text) <> Trim(curDistrict_.ZoneName) And cboZone.ListIndex > 0) Or Trim(txtName.Text) <> Trim(curDistrict_.DistrictName) Or Trim(txtSaleStaff.Text) <> Trim(curDistrict_.SaleStaffName) Then
        isChanged = True
    Else
        isChanged = False
    End If
End Function

Private Sub cboZone_Click()
    tSql = " AND isActive = 'Y'"
    ListDistrict
End Sub

Private Sub ListDistrict()
Dim tmpS As New BWGSearchManager
Dim x%, tmpList As ListItem
    lvList.ListItems.Clear
    If cboZone.ListIndex > 0 Then
        Set colDistrict_ = tmpS.SaleDistrictSearch("SaleZoneID='" & colZone_(cboZone.ListIndex).ID & "'" & tSql)
        For x = 1 To colDistrict_.Count
            Set tmpList = lvList.ListItems.Add(, "'" & colDistrict_(x).ID & "'", colDistrict_(x).DistrictName)
            tmpList.SubItems(1) = colDistrict_(x).SaleStaffName
            tmpList.SubItems(2) = colDistrict_(x).isActive
            If UCase(colDistrict_(x).isActive) = "Y" Then
                tmpList.Checked = True
            Else
                tmpList.Checked = False
            End If
            Set tmpList = Nothing
            colDistrict_(x).isChanged = False
        Next
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
    CreateNewDisctrict
    ClearScreen
    txtName.SetFocus
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

    If curDistrict_ Is Nothing Then Exit Sub
    If cboZone.ListIndex = 0 Then
        MsgBox "กรุณาใส่ชื่อ Zone ก่อนเก็บข้อมูล", vbExclamation
        cboZone.SetFocus
        Exit Sub
    End If
    If Trim(txtName.Text) = "" Then
        MsgBox "กรุณาระบุชื่อเขตก่อนเก็บข้อมูล", vbExclamation
        txtName.SetFocus
        Exit Sub
    End If
    If Trim(txtSaleStaff.Text) = "" Then
        MsgBox "กรุณาระบุชื่อเซลล์ที่รับผิดชอบก่อนเก็บข้อมูล", vbExclamation
        txtSaleStaff.SetFocus
        Exit Sub
    End If

    If isChanged Then
        
        If Trim(curDistrict_.ID) = "" Then
            curDistrict_.ID = GetGUID
            DBConnExc "Insert into tbSaleDistrict (SaleDistrictID,SaleZoneID,SDName,SaleStaffID) Values ('" & Trim(curDistrict_.ID) & "','" & (colZone_(cboZone.ListIndex).ID) & "','" & Trim(txtName.Text) & "','" & Trim(curStaff_.ID) & "')"
            Set tmpList = lvList.ListItems.Add(, "'" & curDistrict_.ID & "'", txtName.Text)
            colDistrict_.Add curDistrict_, "'" & curDistrict_.ID & "'"
        Else
            DBConnExc "Update tbSaleDistrict Set SaleZoneID='" & Trim(colZone_(cboZone.ListIndex).ID) & "',SDName='" & txtName.Text & "',SaleStaffID='" & curStaff_.ID & "' WHERE SaleDistrictID='" & curDistrict_.ID & "'"
            Set tmpList = lvList.ListItems("'" & curDistrict_.ID & "'")
        End If
        tmpList.Text = Trim(txtName.Text)
        tmpList.SubItems(1) = txtSaleStaff.Text
        curDistrict_.DistrictName = Trim(txtName.Text)
        curDistrict_.SaleStaffID = curStaff_.ID
        curDistrict_.SaleStaffName = Trim(curStaff_.StaffFName & " " & curStaff_.StaffLName)
        curDistrict_.ZoneID = colZone_(cboZone.ListIndex).ID
        curDistrict_.ZoneName = cboZone.Text
        Set tmpList = Nothing
        Set tmpCon = Nothing
    End If
    CreateNewDisctrict
    ClearScreen
    txtName.SetFocus
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
    'Set colDistrict_ = searchM.SaleZoneSearch
    'ShowUnitList
    PopulateZone
    CreateNewDisctrict
End Sub

Private Sub ShowDistrictList()
Dim x%
Dim tmpList As ListItem
Dim tmpItem As BWGSaleDistrict
    If colDistrict_ Is Nothing Then Exit Sub
    lvList.ListItems.Clear
    For x = 1 To colDistrict_.Count
        Set tmpItem = colDistrict_(x)
        Set tmpList = lvList.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.DistrictName)
        tmpList.SubItems(1) = tmpItem.SaleStaffName
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
End Sub

Private Sub lvList_DblClick()
    If Not callerForm_ Is Nothing And Not lvList.SelectedItem Is Nothing Then
        callerForm_.SetDataDistrict colDistrict_(lvList.SelectedItem.Key)
        Unload Me
    End If
End Sub

Private Sub lvList_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim tSql As String
    If Item.Checked = False Then
        tSql = "Exec sp_UpdateSaleDistrictStatus " & Item.Key & ", 'N'"
        DBConnExc tSql
    ElseIf Item.Checked = True Then
        tSql = "Exec sp_UpdateSaleDistrictStatus " & Item.Key & ", 'Y'"
        DBConnExc tSql
    End If
End Sub

Private Sub lvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim tmpKey As String
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    If lvList.SelectedItem Is Nothing Then Exit Sub
    Set curDistrict_ = Nothing
    tmpKey = Trim(lvList.SelectedItem.Key)
    If isExist(colDistrict_, tmpKey) Then
        Set curDistrict_ = colDistrict_(tmpKey)
        txtName.Text = curDistrict_.DistrictName
        txtSaleStaff.Text = curDistrict_.SaleStaffName
        Set tmpCol = tmpS.UserSearch("StaffID='" & curDistrict_.SaleStaffID & "'")
        If tmpCol.Count > 0 Then Set curStaff_ = tmpCol(1)
        Set tmpCol = Nothing
        Set tmpS = Nothing
        curRole_ = UpdateRole
    End If
End Sub

Private Sub CreateNewDisctrict()
    Set curDistrict_ = New BWGSaleDistrict
    curDistrict_.ID = ""
    curRole_ = AddNewRole
End Sub

Private Sub ClearScreen()
    txtName.Text = ""
    txtSaleStaff.Text = ""
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
    If curDistrict_.ID = "" Then
        MsgBox "กรุณาเลือกรายการที่ต้องการลบ", vbExclamation
        Exit Sub
    End If
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        tmpKey = "'" & curDistrict_.ID & "'"
        If MsgBox("ต้องการลบข้อมูลรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            tmpStr = curDistrict_.Delete
            If Trim(tmpStr) = "" Then
                If isExist(lvList.ListItems, tmpKey) Then lvList.ListItems.Remove tmpKey
                If isExist(colDistrict_, tmpKey) Then colDistrict_.Remove tmpKey
                CreateNewDisctrict
                ClearScreen
                txtName.SetFocus
            Else
                MsgBox tmpStr, vbExclamation
            End If
        End If
    End If
End Sub

Private Sub PopulateZone()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colZone_ = tmpS.SaleZoneSearch
    cboZone.Clear
    cboZone.AddItem "== กรุณาระบุ Zone =="
    For x = 1 To colZone_.Count
        cboZone.AddItem colZone_(x).ZoneName
    Next
    cboZone.ListIndex = 0
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff)
    Set curStaff_ = selStaff
    With curStaff_
        txtSaleStaff.Text = .StaffFName & " " & .StaffLName
    End With
End Sub

Private Sub lvList_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MNU
    End If
End Sub

Private Sub MNU_Active_Click()
    tSql = " AND isActive = 'Y'"
    ListDistrict
End Sub

Private Sub MNU_All_Click()
    tSql = ""
    ListDistrict
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
