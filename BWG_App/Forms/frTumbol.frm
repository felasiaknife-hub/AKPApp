VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTumbol 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F137"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   750
      MaxLength       =   255
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "เก็บข้อมูล"
      Height          =   795
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5490
      Width           =   1065
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5490
      Width           =   1040
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4245
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   7488
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   11994
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ตำบล"
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
      TabIndex        =   5
      Top             =   0
      Width           =   495
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
      Picture         =   "frTumbol.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ชื่อตำบล"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   570
   End
End
Attribute VB_Name = "frmTumbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curAumphur_ As IMIAumphur
Dim firstActive As Boolean

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim tmpItem As New IMITumbol
Dim tmpLVItem As ListItem
Dim isNew As Boolean
    If Trim(txtDesc.Text) = vbNullString Then
        Set tmpItem = Nothing
        MsgBox "กรุณาป้อนข้อมูลชื่อตำบล", vbExclamation
        txtDesc.SetFocus
        Exit Sub
    End If
    
    Set tmpLVItem = lvList.FindItem(Trim(txtDesc))
    If tmpLVItem Is Nothing Then
        tmpItem.ID = getNewProductID("TB")
        tmpItem.AumphurID = curAumphur_.ID
        tmpItem.ProvinceID = curAumphur_.ProvinceID
        tmpItem.ProvinceName = curAumphur_.ProvinceName
        isNew = True
    Else
        tmpItem.ID = Trim(Replace(tmpLVItem.Key, "'", ""))
        isNew = False
    End If
    tmpItem.Name = Trim(txtDesc)
    tmpItem.Save
    
    If isNew Then
        lvList.ListItems.Add , "'" & tmpItem.ID & "'", tmpItem.Name
    Else
        tmpLVItem.Text = tmpItem.Name
    End If
    
    Set tmpItem = Nothing
    Set tmpLVItem = Nothing
    
    lvList.Refresh
    txtDesc.Text = ""
    txtDesc.SetFocus
End Sub

Private Sub Form_Activate()
Dim tmpCol As Collection
Dim tmpListItem As IMITumbol
Dim tmpS As New BWGSearchManager
Dim x%
    
    If firstActive Then
        Set tmpCol = tmpS.TumbolSearch("AumphurID='" & curAumphur_.ID & "'")
        For x = 1 To tmpCol.Count
            Set tmpListItem = tmpCol(x)
            lvList.ListItems.Add , "'" & tmpListItem.ID & "'", tmpListItem.Name
            Set tmpListItem = Nothing
        Next
        Set tmpS = Nothing
        firstActive = False
    End If
    
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
        cmdSave.Picture = .picSave.Picture
        cmdExit.Picture = .picExit.Picture
    End With
End Sub

Private Sub lvList_ItemClick(ByVal item As MSComctlLib.ListItem)
    If Not item Is Nothing Then
        txtDesc.Text = Trim(item.Text)
    End If
End Sub

Private Sub lvList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpItem As IMITumbol
    If lvList.SelectedItem Is Nothing Then Exit Sub
    
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        If MsgBox("ต้องการลบข้อมูลรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            Set tmpItem = New IMITumbol
            tmpItem.ID = Replace(lvList.SelectedItem.Key, "'", "")
            tmpItem.Delete
            lvList.ListItems.Remove lvList.SelectedItem.Index
            Set tmpItem = Nothing
            txtDesc.Text = ""
            txtDesc.SetFocus
        End If
    End If
End Sub

Private Sub txtDesc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdSave_Click
    End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Public Sub SetCurAumphur(selAumphur As IMIAumphur)
    Set curAumphur_ = selAumphur
End Sub
