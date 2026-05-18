VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEUCode 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   $"frEUCode.frx":0000
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frEUCode.frx":000C
   ScaleHeight     =   9345
   ScaleWidth      =   11520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "EU Code ชุดที่ 3"
      ForeColor       =   &H80000008&
      Height          =   2835
      Index           =   2
      Left            =   30
      TabIndex        =   22
      Top             =   6480
      Width           =   10365
      Begin VB.ComboBox cboHAType 
         Height          =   315
         ItemData        =   "frEUCode.frx":0E16
         Left            =   7080
         List            =   "frEUCode.frx":0E20
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "สร้างใหม่"
         Height          =   315
         Left            =   9330
         TabIndex        =   14
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton Command5 
         Caption         =   "เก็บข้อมูล"
         Height          =   315
         Left            =   8370
         TabIndex        =   13
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox txtCode3Desc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3150
         MaxLength       =   255
         TabIndex        =   11
         Top             =   240
         Width           =   3855
      End
      Begin VB.TextBox txtCode3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         MaxLength       =   255
         TabIndex        =   10
         Top             =   240
         Width           =   1485
      End
      Begin MSComctlLib.ListView lvCode3 
         Height          =   2145
         Left            =   60
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   630
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   3784
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
            Text            =   "Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "คำอธิบาย"
            Object.Width           =   12524
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "HA/HN"
            Object.Width           =   1587
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "คำอธิบาย"
         Height          =   195
         Index           =   6
         Left            =   2430
         TabIndex        =   26
         Top             =   300
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Code"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   23
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "EU Code ชุดที่ 2"
      ForeColor       =   &H80000008&
      Height          =   2835
      Index           =   1
      Left            =   30
      TabIndex        =   20
      Top             =   3600
      Width           =   11475
      Begin VB.CommandButton Command4 
         Caption         =   "สร้างใหม่"
         Height          =   315
         Left            =   10440
         TabIndex        =   8
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton Command3 
         Caption         =   "เก็บข้อมูล"
         Height          =   315
         Left            =   9480
         TabIndex        =   7
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox txtCode2Desc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3150
         MaxLength       =   255
         TabIndex        =   6
         Top             =   240
         Width           =   6285
      End
      Begin VB.TextBox txtCode2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         MaxLength       =   255
         TabIndex        =   5
         Top             =   240
         Width           =   1485
      End
      Begin MSComctlLib.ListView lvCode2 
         Height          =   2145
         Left            =   60
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   630
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   3784
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
            Text            =   "Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "คำอธิบาย"
            Object.Width           =   15875
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "คำอธิบาย"
         Height          =   195
         Index           =   5
         Left            =   2430
         TabIndex        =   25
         Top             =   300
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Code"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   21
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "EU Code ชุดที่ 1"
      ForeColor       =   &H80000008&
      Height          =   2835
      Index           =   0
      Left            =   30
      TabIndex        =   18
      Top             =   720
      Width           =   11475
      Begin VB.CommandButton Command2 
         Caption         =   "สร้างใหม่"
         Height          =   315
         Left            =   10440
         TabIndex        =   3
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton Command1 
         Caption         =   "เก็บข้อมูล"
         Height          =   315
         Left            =   9480
         TabIndex        =   2
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox txtCode1Desc 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3150
         MaxLength       =   255
         TabIndex        =   1
         Top             =   240
         Width           =   6285
      End
      Begin VB.TextBox txtCode1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   720
         MaxLength       =   255
         TabIndex        =   0
         Top             =   240
         Width           =   1485
      End
      Begin MSComctlLib.ListView lvCode1 
         Height          =   2145
         Left            =   60
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   630
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   3784
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
            Text            =   "Code"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "คำอธิบาย"
            Object.Width           =   15875
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "คำอธิบาย"
         Height          =   195
         Index           =   4
         Left            =   2430
         TabIndex        =   24
         Top             =   300
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Code"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   19
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8520
      Width           =   1040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ข้อมูล EU Code"
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
      TabIndex        =   17
      Top             =   0
      Width           =   1365
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
      Picture         =   "frEUCode.frx":0E2C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11550
   End
End
Attribute VB_Name = "frmEUCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim firstActive As Boolean
Dim curEUCode1_ As BWGEUCode1
Dim curEUCode2_ As BWGEUCode2
Dim curEUCode3_ As BWGEUCode3

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub SaveEUCode1()
Dim tmpItem As New BWGEUCode1
Dim tmpLVItem As ListItem
Dim isNew As Boolean
    If Trim(txtCode1.Text) = vbNullString Then
        Set tmpItem = Nothing
        MsgBox "กรุณาป้อนข้อมูลชื่อจังหวัด", vbExclamation
        txtCode1.SetFocus
        Exit Sub
    End If
    
    Set tmpLVItem = lvCode1.FindItem(Trim(txtCode1.Text))
    If Not curEUCode1_ Is Nothing And Not tmpLVItem Is Nothing Then
        If tmpLVItem.Key <> "'" & curEUCode1_.ID & "'" Then
            MsgBox "Code ซ้ำกับ Code เดิมที่มีอยู่ในระบบ ไม่สามารถจัดเก็บข้อมูลได้", vbExclamation
            txtCode1.SetFocus
            Exit Sub
        End If
    Else
        If Not tmpLVItem Is Nothing Then
            MsgBox "Code ซ้ำกับ Code เดิมที่มีอยู่ในระบบ ไม่สามารถจัดเก็บข้อมูลได้", vbExclamation
            txtCode1.SetFocus
            Exit Sub
        End If
    End If
    If isExist(lvCode1.ListItems, "'" & curEUCode1_.ID & "'") Then
        Set tmpLVItem = lvCode1.ListItems("'" & curEUCode1_.ID & "'")
    Else
        Set tmpLVItem = Nothing
    End If
    If tmpLVItem Is Nothing Then
        tmpItem.ID = GetGUID
        isNew = True
    Else
        tmpItem.ID = Trim(Replace(tmpLVItem.Key, "'", ""))
        isNew = False
    End If
    tmpItem.Code1 = Trim(txtCode1)
    tmpItem.CodeDesc = Trim(txtCode1Desc.Text)
    tmpItem.Save
    
    If isNew Then
        Set tmpLVItem = lvCode1.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.Code1)
    Else
        tmpLVItem.Text = tmpItem.Code1
    End If
    tmpLVItem.SubItems(1) = Trim(txtCode1Desc.Text)
    
    Set tmpItem = Nothing
    Set tmpLVItem = Nothing
    
    lvCode1.Refresh
    Command2_Click
End Sub

Private Sub Command1_Click()
    If Trim(txtCode1.Text) = "" Then
        MsgBox "กรุณาใส่ข้อมูล Code", vbExclamation
        txtCode1.SetFocus
        Exit Sub
    End If
    SaveEUCode1
End Sub

Private Sub Command2_Click()
    Set curEUCode1_ = Nothing
    Set curEUCode1_ = New BWGEUCode1
    curEUCode1_.ID = GetGUID
    txtCode1.Text = ""
    txtCode1Desc.Text = ""
    txtCode1.SetFocus
End Sub

Private Sub Command3_Click()
    If Trim(txtCode2.Text) = "" Then
        MsgBox "กรุณาใส่ข้อมูล Code", vbExclamation
        txtCode2.SetFocus
        Exit Sub
    End If
    SaveEUCode2
End Sub

Private Sub Command4_Click()
    Set curEUCode2_ = Nothing
    Set curEUCode2_ = New BWGEUCode2
    curEUCode2_.ID = GetGUID
    txtCode2.Text = ""
    txtCode2Desc.Text = ""
    txtCode2.SetFocus
End Sub

Private Sub Command5_Click()
    If Trim(txtCode3.Text) = "" Then
        MsgBox "กรุณาใส่ข้อมูล Code", vbExclamation
        txtCode3.SetFocus
        Exit Sub
    End If
    SaveEUCode3
End Sub

Private Sub Command6_Click()
    Set curEUCode3_ = Nothing
    Set curEUCode3_ = New BWGEUCode3
    curEUCode3_.ID = GetGUID
    txtCode3.Text = ""
    txtCode3Desc.Text = ""
    cboHAType.Text = ""
    txtCode3.SetFocus
End Sub

Private Sub Form_Activate()
Dim tmpCol As Collection
Dim tmpListItem As BWGEUCode1
Dim tmpS As New BWGSearchManager
Dim tmpItem As ListItem
Dim x%
   
    If firstActive Then
        Set tmpCol = tmpS.EUCode1Search
        For x = 1 To tmpCol.Count
            Set tmpListItem = tmpCol(x)
            Set tmpItem = lvCode1.ListItems.Add(, "'" & tmpListItem.ID & "'", tmpListItem.Code1)
            tmpItem.SubItems(1) = tmpListItem.CodeDesc
            Set tmpItem = Nothing
            Set tmpListItem = Nothing
        Next
        Set tmpS = Nothing
        firstActive = False
        SetEUCode2Visible False
        SetEUCode3Visible False

    End If
    
    Set tmpCol = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    firstActive = True
    With frmPicture
        cmdExit.Picture = .picExit.Picture
    End With
    Set curEUCode1_ = New BWGEUCode1
    Set curEUCode2_ = New BWGEUCode2
    Set curEUCode3_ = New BWGEUCode3

End Sub

Private Sub lvCode2_ItemClick(ByVal item As MSComctlLib.ListItem)
Dim tmpS As BWGSearchManager
Dim tmpCol As Collection
Dim x%
Dim tmpList As ListItem
    If Not item Is Nothing Then
        If curEUCode2_ Is Nothing Then Set curEUCode2_ = New BWGEUCode2
        curEUCode2_.ID = Replace(item.Key, "'", "")
        curEUCode2_.EUCode1ID = Trim(curEUCode1_.ID)
        'curEUCode2_.Code1 = Trim(curEUCode1_.Code1)
        curEUCode2_.Code2 = Trim(item.Text)
        curEUCode2_.CodeDesc = item.SubItems(1)
        Set curEUCode3_ = Nothing
        Set curEUCode3_ = New BWGEUCode3
        Frame1(2).Caption = "EU Code ชุดที่ 3 ของ " & curEUCode1_.Code1 & "-" & curEUCode2_.Code2
        txtCode2.Text = Trim(item.Text)
        txtCode2Desc.Text = item.SubItems(1)
        Set tmpS = New BWGSearchManager
        Set tmpCol = tmpS.EUCode3Search("EUCode2ID=" & item.Key)
        lvCode3.ListItems.Clear
        For x = 1 To tmpCol.Count
            Set tmpList = lvCode3.ListItems.Add(, "'" & tmpCol(x).ID & "'", tmpCol(x).Code3)
            tmpList.SubItems(1) = tmpCol(x).CodeDesc
            tmpList.SubItems(2) = tmpCol(x).HAType
            Set tmpList = Nothing
        Next
        Set tmpCol = Nothing
        Set tmpS = Nothing
        SetEUCode3Visible True
    End If
End Sub

Private Sub lvCode2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpItem As BWGEUCode2
Dim tmpStr As String
    If lvCode2.SelectedItem Is Nothing Then Exit Sub
    
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        If MsgBox("ต้องการลบข้อมูล EU Code ชุดที่ 2 นี้ออกจากฐานข้อมูลหรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            Set tmpItem = New BWGEUCode2
            tmpItem.ID = Replace(lvCode2.SelectedItem.Key, "'", "")
            tmpStr = Trim(tmpItem.Delete)
            If Trim(tmpStr) <> "" Then
                MsgBox "ไม่สามารถลบข้อมูลได้ เนื่องจาก " & vbCrLf & tmpStr, vbExclamation
                Exit Sub
            End If
            lvCode2.ListItems.Remove lvCode2.SelectedItem.Index
            Set tmpItem = Nothing
            txtCode2.Text = ""
            txtCode2.SetFocus
            Set curEUCode2_ = Nothing
            
            txtCode3.Text = ""
            lvCode3.ListItems.Clear
            Set curEUCode3_ = Nothing
            SetEUCode3Visible False
        End If
    End If
End Sub

Private Sub lvCode1_ItemClick(ByVal item As MSComctlLib.ListItem)
Dim tmpS As BWGSearchManager
Dim tmpCol As Collection
Dim x%
Dim tmpList As ListItem
    If Not item Is Nothing Then
        If curEUCode1_ Is Nothing Then Set curEUCode1_ = New BWGEUCode1
        curEUCode1_.ID = Replace(item.Key, "'", "")
        curEUCode1_.Code1 = Trim(item.Text)
        curEUCode1_.CodeDesc = item.SubItems(1)
        Set curEUCode2_ = Nothing
        Set curEUCode2_ = New BWGEUCode2
        Set curEUCode3_ = Nothing
        
        txtCode1.Text = Trim(item.Text)
        txtCode1Desc.Text = item.SubItems(1)
        Set tmpS = New BWGSearchManager
        Set tmpCol = tmpS.EUCode2Search("EUCode1ID=" & item.Key)
        lvCode2.ListItems.Clear
        txtCode2.Text = ""
        txtCode2Desc.Text = ""
        lvCode3.ListItems.Clear
        txtCode3.Text = ""
        txtCode3Desc.Text = ""
        cboHAType.Text = ""
        Frame1(1).Caption = "EU Code ชุดที่ 2 ของ " & item.Text
        For x = 1 To tmpCol.Count
            Set tmpList = lvCode2.ListItems.Add(, "'" & tmpCol(x).ID & "'", tmpCol(x).Code2)
            tmpList.SubItems(1) = tmpCol(x).CodeDesc
            Set tmpList = Nothing
        Next
        Set tmpCol = Nothing
        Set tmpS = Nothing
        SetEUCode2Visible True
        SetEUCode3Visible False
    End If
End Sub

Private Sub lvCode1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpItem As BWGEUCode1
Dim tmpStr As String
    If lvCode1.SelectedItem Is Nothing Then Exit Sub
    
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        If MsgBox("ต้องการลบข้อมูล EU Code ชุดที่ 1 นี้ออกจากฐานข้อมูลหรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            Set tmpItem = New BWGEUCode1
            tmpItem.ID = Replace(lvCode1.SelectedItem.Key, "'", "")
            tmpStr = Trim(tmpItem.Delete)
            If Trim(tmpStr) <> "" Then
                MsgBox "ไม่สามารถลบข้อมูลได้ เนื่องจาก " & vbCrLf & tmpStr, vbExclamation
                Exit Sub
            End If
            lvCode1.ListItems.Remove lvCode1.SelectedItem.Index
            Set tmpItem = Nothing
            txtCode1.Text = ""
            txtCode1.SetFocus
            Set curEUCode1_ = Nothing
            
            txtCode2.Text = ""
            lvCode2.ListItems.Clear
            Set curEUCode2_ = Nothing
            
            txtCode3.Text = ""
            lvCode3.ListItems.Clear
            Set curEUCode3_ = Nothing
            
            SetEUCode2Visible False
            SetEUCode3Visible False
        End If
    End If
End Sub

Private Sub lvCode3_ItemClick(ByVal item As MSComctlLib.ListItem)
    If Not item Is Nothing Then
        If curEUCode3_ Is Nothing Then Set curEUCode3_ = New BWGEUCode3
        curEUCode3_.ID = Replace(item.Key, "'", "")
        curEUCode3_.EUCode2ID = Trim(curEUCode2_.ID)
        'curEUCode3_.AumphurName = Trim(curEUCode2_.Name)
        curEUCode3_.Code3 = Trim(item.Text)
        curEUCode3_.CodeDesc = item.SubItems(1)
        
        txtCode3.Text = Trim(item.Text)
        txtCode3Desc.Text = item.SubItems(1)
        cboHAType.Text = item.SubItems(2)
    End If
End Sub

Private Sub lvCode3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpItem As BWGEUCode3
Dim tmpStr As String
    If lvCode3.SelectedItem Is Nothing Then Exit Sub
    
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        If MsgBox("ต้องการลบข้อมูล EU Code ชุดที่ 3 นี้ออกจากฐานข้อมูลหรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            Set tmpItem = New BWGEUCode3
            tmpItem.ID = Replace(lvCode3.SelectedItem.Key, "'", "")
            tmpStr = Trim(tmpItem.Delete)
            If Trim(tmpStr) <> "" Then
                MsgBox "ไม่สามารถลบข้อมูลได้ เนื่องจาก " & vbCrLf & tmpStr, vbExclamation
                Exit Sub
            End If
            lvCode3.ListItems.Remove lvCode3.SelectedItem.Index
            Set tmpItem = Nothing
            txtCode3.Text = ""
            Set curEUCode3_ = Nothing
        End If
    End If
End Sub

Private Sub txtCode2_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        SaveEUCode2
'    End If
End Sub

Private Sub txtCode1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        SaveEUCode1
'    End If
End Sub

Private Sub txtCode1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCode3_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        SaveEUCode3
'    End If
End Sub

Private Sub SaveEUCode2()
Dim tmpItem As New BWGEUCode2
Dim tmpLVItem As ListItem
Dim isNew As Boolean
    If Trim(txtCode2.Text) = vbNullString Then
        Set tmpItem = Nothing
        MsgBox "กรุณาป้อนข้อมูล Code", vbExclamation
        txtCode2.SetFocus
        Exit Sub
    End If
    
    Set tmpLVItem = lvCode2.FindItem(Trim(txtCode2.Text))
    If Not curEUCode2_ Is Nothing And Not tmpLVItem Is Nothing Then
        If tmpLVItem.Key <> "'" & curEUCode2_.ID & "'" Then
            MsgBox "Code ซ้ำกับ Code เดิมที่มีอยู่ในระบบ ไม่สามารถจัดเก็บข้อมูลได้", vbExclamation
            txtCode2.SetFocus
            Exit Sub
        End If
    Else
        If Not tmpLVItem Is Nothing Then
            MsgBox "Code ซ้ำกับ Code เดิมที่มีอยู่ในระบบ ไม่สามารถจัดเก็บข้อมูลได้", vbExclamation
            txtCode2.SetFocus
            Exit Sub
        End If
    End If
    If isExist(lvCode2.ListItems, "'" & curEUCode2_.ID & "'") Then
        Set tmpLVItem = lvCode2.ListItems("'" & curEUCode2_.ID & "'")
    Else
        Set tmpLVItem = Nothing
    End If
    If isExist(lvCode2.ListItems, "'" & curEUCode2_.ID & "'") Then
        Set tmpLVItem = lvCode2.ListItems("'" & curEUCode2_.ID & "'")
    Else
        Set tmpLVItem = Nothing
    End If
    
    If tmpLVItem Is Nothing Then
        tmpItem.ID = GetGUID
        isNew = True
    Else
        tmpItem.ID = Trim(Replace(tmpLVItem.Key, "'", ""))
        isNew = False
    End If
    tmpItem.Code2 = Trim(txtCode2.Text)
    tmpItem.EUCode1ID = Trim(curEUCode1_.ID)
    tmpItem.CodeDesc = Trim(txtCode2Desc.Text)
    tmpItem.Save
    
    If isNew Then
        Set tmpLVItem = lvCode2.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.Code2)
    Else
        tmpLVItem.Text = tmpItem.Code2
    End If
    tmpLVItem.SubItems(1) = Trim(txtCode2Desc.Text)
    Set tmpItem = Nothing
    Set tmpLVItem = Nothing
    
    lvCode2.Refresh
    Command4_Click
End Sub

Private Sub SaveEUCode3()
Dim tmpItem As New BWGEUCode3
Dim tmpLVItem As ListItem
Dim isNew As Boolean
    If Trim(txtCode3.Text) = vbNullString Then
        Set tmpItem = Nothing
        MsgBox "กรุณาป้อนข้อมูล Code", vbExclamation
        txtCode3.SetFocus
        Exit Sub
    End If
    
    Set tmpLVItem = lvCode3.FindItem(Trim(txtCode3.Text))
    If Not curEUCode3_ Is Nothing And Not tmpLVItem Is Nothing Then
        If tmpLVItem.Key <> "'" & curEUCode3_.ID & "'" Then
            MsgBox "Code ซ้ำกับ Code เดิมที่มีอยู่ในระบบ ไม่สามารถจัดเก็บข้อมูลได้", vbExclamation
            txtCode3.SetFocus
            Exit Sub
        End If
    Else
        If Not tmpLVItem Is Nothing Then
            MsgBox "Code ซ้ำกับ Code เดิมที่มีอยู่ในระบบ ไม่สามารถจัดเก็บข้อมูลได้", vbExclamation
            txtCode3.SetFocus
            Exit Sub
        End If
    End If
    If isExist(lvCode3.ListItems, "'" & curEUCode3_.ID & "'") Then
        Set tmpLVItem = lvCode3.ListItems("'" & curEUCode3_.ID & "'")
    Else
        Set tmpLVItem = Nothing
    End If
    If isExist(lvCode3.ListItems, "'" & curEUCode3_.ID & "'") Then
        Set tmpLVItem = lvCode3.ListItems("'" & curEUCode3_.ID & "'")
    Else
        Set tmpLVItem = Nothing
    End If
    
    If tmpLVItem Is Nothing Then
        tmpItem.ID = GetGUID
        isNew = True
    Else
        tmpItem.ID = Trim(Replace(tmpLVItem.Key, "'", ""))
        isNew = False
    End If
    tmpItem.Code3 = Trim(txtCode3.Text)
    tmpItem.EUCode2ID = Trim(curEUCode2_.ID)
    tmpItem.CodeDesc = Trim(txtCode3Desc.Text)
    tmpItem.HAType = Trim(cboHAType.Text)
    tmpItem.Save
    
    If isNew Then
        Set tmpLVItem = lvCode3.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.Code3)
    Else
        tmpLVItem.Text = tmpItem.Code3
    End If
    tmpLVItem.SubItems(1) = Trim(txtCode3Desc.Text)
    tmpLVItem.SubItems(2) = Trim(cboHAType.Text)
    Set tmpItem = Nothing
    Set tmpLVItem = Nothing
    
    lvCode3.Refresh
    Command6_Click
End Sub

Private Sub SetEUCode2Visible(vType As Boolean)
    Label1(2).Visible = vType
    Label1(5).Visible = vType
    txtCode2.Visible = vType
    txtCode2Desc.Visible = vType
    lvCode2.Visible = vType
    Command3.Visible = vType
    Command4.Visible = vType
    If Not vType Then Frame1(1).Caption = "EU Code ชุดที่ 2"
End Sub

Private Sub SetEUCode3Visible(vType As Boolean)
    Label1(3).Visible = vType
    Label1(6).Visible = vType
    txtCode3.Visible = vType
    txtCode3Desc.Visible = vType
    cboHAType.Visible = vType
    lvCode3.Visible = vType
    Command5.Visible = vType
    Command6.Visible = vType
    If Not vType Then Frame1(2).Caption = "EU Code ชุดที่ 3"
End Sub
