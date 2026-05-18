VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProvince 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F109"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCountry 
      Height          =   315
      Left            =   1020
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   720
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ตำบล"
      ForeColor       =   &H80000008&
      Height          =   5685
      Index           =   2
      Left            =   7650
      TabIndex        =   12
      Top             =   1080
      Width           =   3825
      Begin VB.TextBox txtTumbol 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   990
         MaxLength       =   255
         TabIndex        =   4
         Top             =   240
         Width           =   2775
      End
      Begin MSComctlLib.ListView lvTumbol 
         Height          =   4995
         Left            =   60
         TabIndex        =   5
         Top             =   630
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   8811
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ชื่อจังหวัด"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "ชื่อตำบล"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   13
         Top             =   300
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "อำเภอ"
      ForeColor       =   &H80000008&
      Height          =   5685
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      Top             =   1080
      Width           =   3825
      Begin VB.TextBox txtAumphur 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   990
         MaxLength       =   255
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin MSComctlLib.ListView lvAumphur 
         Height          =   4995
         Left            =   60
         TabIndex        =   3
         Top             =   630
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   8811
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ชื่อจังหวัด"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "ชื่ออำเภอ"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   11
         Top             =   300
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "จังหวัด"
      ForeColor       =   &H80000008&
      Height          =   5685
      Index           =   0
      Left            =   30
      TabIndex        =   8
      Top             =   1080
      Width           =   3825
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   990
         MaxLength       =   255
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin MSComctlLib.ListView lvProvince 
         Height          =   4995
         Left            =   60
         TabIndex        =   1
         Top             =   630
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   8811
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ชื่อจังหวัด"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "ชื่อจังหวัด"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   9
         Top             =   300
         Width           =   660
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6810
      Width           =   1040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ประเทศ"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   780
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ข้อมูล จังหวัด/อำเภอ/ตำบล"
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
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   2340
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   90
      X2              =   6060
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   5910
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frProvince.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11550
   End
End
Attribute VB_Name = "frmProvince"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim firstActive As Boolean
Dim curProvince_ As IMIProvince
Dim curAumphur_ As IMIAumphur
Dim curTumbol_ As IMITumbol
Dim tSql As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub SaveProvince()
Dim tmpItem As New IMIProvince
Dim tmpLVItem As ListItem
Dim isNew As Boolean
Dim rs As New ADODB.Recordset

    If Trim(TxtName.Text) = vbNullString Then
        Set tmpItem = Nothing
        MsgBox "กรุณาป้อนข้อมูลชื่อจังหวัด", vbExclamation
        TxtName.SetFocus
        Exit Sub
    End If
    
    Set rs = GetRS("SELECT * FROM THAILAND.dbo.province WHERE PROVINCE_NAME = '" & Trim(TxtName.Text) & "'", adOpenForwardOnly, adLockReadOnly)
    If rs.RecordCount <= 0 Then
        MsgBox "ไม่มีจังหวัด """ & Trim(TxtName.Text) & """ ในประเทศ " & cboCountry.Text
        Set tmpLVItem = Nothing
        TxtName.Text = ""
        Exit Sub
    End If
    
    Set tmpLVItem = lvProvince.FindItem(Trim(TxtName))
    If tmpLVItem Is Nothing Then
        tmpItem.ID = getNewProductID("PRV")
        isNew = True
    Else
        tmpItem.ID = Trim(Replace(tmpLVItem.Key, "'", ""))
        isNew = False
    End If
    tmpItem.ProvinceName = Trim(TxtName)
    tmpItem.Save
    
    If isNew Then
        Set tmpLVItem = lvProvince.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.ProvinceName)
    Else
        tmpLVItem.Text = tmpItem.ProvinceName
    End If
    
    Set tmpItem = Nothing
    Set tmpLVItem = Nothing
    
    lvProvince.Refresh
    TxtName.Text = ""
    TxtName.SetFocus
End Sub

Private Sub Form_Activate()
Dim tmpCol As Collection
Dim tmpListItem As IMIProvince
Dim tmpS As New BWGSearchManager
Dim tmpItem As ListItem
Dim x%
   
    If firstActive Then
        Set tmpCol = tmpS.ProvinceSearch
        For x = 1 To tmpCol.Count
            Set tmpListItem = tmpCol(x)
            Set tmpItem = lvProvince.ListItems.Add(, "'" & tmpListItem.ID & "'", tmpListItem.ProvinceName)
            Set tmpItem = Nothing
            Set tmpListItem = Nothing
        Next
        Set tmpS = Nothing
        firstActive = False
        SetAumphurVisible False
        SetTumbolVisible False

    End If
    
    cboCountry.AddItem "THAILAND"
    cboCountry.ListIndex = 0
    
    Set tmpCol = Nothing
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    firstActive = True
    With frmPicture
        cmdExit.Picture = .picExit.Picture
    End With
End Sub

Private Sub lvAumphur_ItemClick(ByVal item As MSComctlLib.ListItem)
Dim tmpS As BWGSearchManager
Dim tmpCol As Collection
Dim x%
    If Not item Is Nothing Then
        If curAumphur_ Is Nothing Then Set curAumphur_ = New IMIAumphur
        curAumphur_.ID = Replace(item.Key, "'", "")
        curAumphur_.ProvinceID = Trim(curProvince_.ID)
        curAumphur_.ProvinceName = Trim(curProvince_.ProvinceName)
        curAumphur_.Name = Trim(item.Text)
        Set curTumbol_ = Nothing
        Frame1(2).Caption = "ตำบลในอำเภอ " & curAumphur_.Name & " จังหวัด " & curAumphur_.ProvinceName
        txtAumphur.Text = Trim(item.Text)
        Set tmpS = New BWGSearchManager
        Set tmpCol = tmpS.TumbolSearch("AumphurID=" & item.Key)
        lvTumbol.ListItems.Clear
        For x = 1 To tmpCol.Count
            lvTumbol.ListItems.Add , "'" & tmpCol(x).ID & "'", tmpCol(x).Name
        Next
        Set tmpCol = Nothing
        Set tmpS = Nothing
        SetTumbolVisible True
    End If
End Sub

Private Sub lvAumphur_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpItem As IMIAumphur
Dim tmpStr As String
    If lvAumphur.SelectedItem Is Nothing Then Exit Sub
    
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        If MsgBox("ต้องการลบข้อมูลอำเภอนี้ออกจากฐานข้อมูลหรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            Set tmpItem = New IMIAumphur
            tmpItem.ID = Replace(lvAumphur.SelectedItem.Key, "'", "")
            tmpStr = Trim(tmpItem.Delete)
            If Trim(tmpStr) <> "" Then
                MsgBox "ไม่สามารถลบข้อมูลได้ เนื่องจาก " & vbCrLf & tmpStr, vbExclamation
                Exit Sub
            End If
            lvAumphur.ListItems.Remove lvAumphur.SelectedItem.Index
            Set tmpItem = Nothing
            txtAumphur.Text = ""
            txtAumphur.SetFocus
            Set curAumphur_ = Nothing
            
            txtTumbol.Text = ""
            lvTumbol.ListItems.Clear
            Set curTumbol_ = Nothing
            SetTumbolVisible False
        End If
    End If
End Sub

Private Sub lvprovince_ItemClick(ByVal item As MSComctlLib.ListItem)
Dim tmpS As BWGSearchManager
Dim tmpCol As Collection
Dim x%
    If Not item Is Nothing Then
        If curProvince_ Is Nothing Then Set curProvince_ = New IMIProvince
        curProvince_.ID = Replace(item.Key, "'", "")
        curProvince_.ProvinceName = Trim(item.Text)
        Set curAumphur_ = Nothing
        Set curTumbol_ = Nothing
        
        TxtName.Text = Trim(item.Text)
        Set tmpS = New BWGSearchManager
        Set tmpCol = tmpS.AumphurSearch("ProvinceID=" & item.Key)
        lvAumphur.ListItems.Clear
        txtAumphur.Text = ""
        lvTumbol.ListItems.Clear
        txtTumbol.Text = ""
        Frame1(1).Caption = "อำเภอในจังหวัด " & item.Text
        For x = 1 To tmpCol.Count
            lvAumphur.ListItems.Add , "'" & tmpCol(x).ID & "'", tmpCol(x).Name
        Next
        Set tmpCol = Nothing
        Set tmpS = Nothing
        SetAumphurVisible True
        SetTumbolVisible False
    End If
End Sub

Private Sub lvprovince_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpItem As IMIProvince
Dim tmpStr As String
    If lvProvince.SelectedItem Is Nothing Then Exit Sub
    
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        If MsgBox("ต้องการลบข้อมูลจังหวัดนี้ออกจากฐานข้อมูลหรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            Set tmpItem = New IMIProvince
            tmpItem.ID = Replace(lvProvince.SelectedItem.Key, "'", "")
            tmpStr = Trim(tmpItem.Delete)
            If Trim(tmpStr) <> "" Then
                MsgBox "ไม่สามารถลบข้อมูลได้ เนื่องจาก " & vbCrLf & tmpStr, vbExclamation
                Exit Sub
            End If
            lvProvince.ListItems.Remove lvProvince.SelectedItem.Index
            Set tmpItem = Nothing
            TxtName.Text = ""
            TxtName.SetFocus
            Set curProvince_ = Nothing
            
            txtAumphur.Text = ""
            lvAumphur.ListItems.Clear
            Set curAumphur_ = Nothing
            
            txtTumbol.Text = ""
            lvTumbol.ListItems.Clear
            Set curTumbol_ = Nothing
            
            SetAumphurVisible False
            SetTumbolVisible False
        End If
    End If
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SaveProvince
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub lvTumbol_ItemClick(ByVal item As MSComctlLib.ListItem)
    If Not item Is Nothing Then
        If curTumbol_ Is Nothing Then Set curTumbol_ = New IMITumbol
        curTumbol_.ID = Replace(item.Key, "'", "")
        curTumbol_.AumphurID = Trim(curAumphur_.ID)
        curTumbol_.AumphurName = Trim(curAumphur_.Name)
        curTumbol_.Name = Trim(item.Text)
        
        txtTumbol.Text = Trim(item.Text)
    End If
End Sub

Private Sub lvTumbol_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpItem As IMITumbol
Dim tmpStr As String
    If lvTumbol.SelectedItem Is Nothing Then Exit Sub
    
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        If MsgBox("ต้องการลบข้อมูลตำบลนี้ออกจากฐานข้อมูลหรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            Set tmpItem = New IMITumbol
            tmpItem.ID = Replace(lvTumbol.SelectedItem.Key, "'", "")
            tmpStr = Trim(tmpItem.Delete)
            If Trim(tmpStr) <> "" Then
                MsgBox "ไม่สามารถลบข้อมูลได้ เนื่องจาก " & vbCrLf & tmpStr, vbExclamation
                Exit Sub
            End If
            lvTumbol.ListItems.Remove lvTumbol.SelectedItem.Index
            Set tmpItem = Nothing
            txtTumbol.Text = ""
            Set curTumbol_ = Nothing
        End If
    End If
End Sub

Private Sub txtAumphur_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SaveAumphur
    End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SaveProvince
    End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTumbol_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SaveTumbol
    End If
End Sub

Private Sub SaveAumphur()
Dim tmpItem As New IMIAumphur
Dim tmpLVItem As ListItem
Dim isNew As Boolean
Dim rs As New ADODB.Recordset

    If Trim(txtAumphur.Text) = vbNullString Then
        Set tmpItem = Nothing
        MsgBox "กรุณาป้อนข้อมูลชื่ออำเภอ", vbExclamation
        txtAumphur.SetFocus
        Exit Sub
    End If
    tSql = "SELECT * FROM THAILAND.dbo.tbamphur WHERE AMPHUR_NAME LIKE '%" & Trim(txtAumphur.Text) & "%'" & vbCrLf
    tSql = tSql & "AND PROVINCE_ID IN (SELECT PROVINCE_ID FROM THAILAND.dbo.province WHERE PROVINCE_NAME = '" & TxtName.Text & "')"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If rs.RecordCount <= 0 Then
        MsgBox "ไม่มีอำเภอ """ & Trim(txtAumphur.Text) & """ ในจังหวัด " & TxtName.Text
        Set tmpLVItem = Nothing
        txtAumphur.Text = ""
        Exit Sub
    Else
        txtAumphur.Text = Trim("" & rs!AMPHUR_NAME)
    End If
    Set rs = Nothing
    
    Set tmpLVItem = lvAumphur.FindItem(Trim(txtAumphur.Text))
    If tmpLVItem Is Nothing Then
        tmpItem.ID = getNewProductID("AP")
        isNew = True
    Else
        tmpItem.ID = Trim(Replace(tmpLVItem.Key, "'", ""))
        isNew = False
    End If
    tmpItem.Name = Trim(txtAumphur.Text)
    tmpItem.ProvinceName = Trim(curProvince_.ProvinceName)
    tmpItem.ProvinceID = Trim(curProvince_.ID)
    tmpItem.Save
    
    If isNew Then
        Set tmpLVItem = lvAumphur.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.Name)
    Else
        tmpLVItem.Text = tmpItem.Name
    End If
    
    Set tmpItem = Nothing
    Set tmpLVItem = Nothing
    
    lvProvince.Refresh
    txtAumphur.Text = ""
    txtAumphur.SetFocus
End Sub

Private Sub SaveTumbol()
Dim tmpItem As New IMITumbol
Dim tmpLVItem As ListItem
Dim isNew As Boolean
Dim rs As New ADODB.Recordset

    If Trim(txtAumphur.Text) = vbNullString Then
        Set tmpItem = Nothing
        MsgBox "กรุณาป้อนข้อมูลชื่อตำบล", vbExclamation
        txtTumbol.SetFocus
        Exit Sub
    End If
    
    tSql = "SELECT * FROM THAILAND.dbo.district WHERE DISTRICT_NAME LIKE '%" & Trim(txtTumbol.Text) & "%'" & vbCrLf
    tSql = tSql & "AND AMPHUR_ID IN (SELECT AMPHUR_ID FROM THAILAND.dbo.amphur WHERE AMPHUR_NAME = '" & txtAumphur.Text & "')"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If rs.RecordCount <= 0 Then
        MsgBox "ไม่มีตำบล """ & Trim(txtTumbol.Text) & """ ในอำเภอ " & txtAumphur.Text
        Set tmpLVItem = Nothing
        txtTumbol.Text = ""
        Exit Sub
    Else
        txtTumbol.Text = Trim("" & rs!DISTRICT_NAME)
    End If
    
    Set tmpLVItem = lvTumbol.FindItem(Trim(txtTumbol.Text))
    If tmpLVItem Is Nothing Then
        tmpItem.ID = getNewProductID("TB")
        isNew = True
    Else
        tmpItem.ID = Trim(Replace(tmpLVItem.Key, "'", ""))
        isNew = False
    End If
    tmpItem.Name = Trim(txtTumbol.Text)
    tmpItem.AumphurName = Trim(curAumphur_.Name)
    tmpItem.AumphurID = Trim(curAumphur_.ID)
    tmpItem.Save
    
    If isNew Then
        Set tmpLVItem = lvTumbol.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.Name)
    Else
        tmpLVItem.Text = tmpItem.Name
    End If
    
    Set tmpItem = Nothing
    Set tmpLVItem = Nothing
    
    lvTumbol.Refresh
    txtTumbol.Text = ""
    txtTumbol.SetFocus
End Sub

Private Sub SetAumphurVisible(vType As Boolean)
    Label1(2).Visible = vType
    txtAumphur.Visible = vType
    lvAumphur.Visible = vType
    If Not vType Then Frame1(1).Caption = "อำเภอ"
End Sub

Private Sub SetTumbolVisible(vType As Boolean)
    Label1(3).Visible = vType
    txtTumbol.Visible = vType
    lvTumbol.Visible = vType
    If Not vType Then Frame1(2).Caption = "ตำบล"
End Sub
