VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAumphurSearch 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "F003"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   30
      Picture         =   "frAumphurSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   915
   End
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Text            =   "*"
      Top             =   960
      Width           =   4305
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   915
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3645
      Left            =   0
      TabIndex        =   1
      Top             =   1950
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   6429
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
         Text            =   "จังหวัด"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "อำเภอ"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label AumphurSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหารายการอำเภอ"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   1890
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   90
      X2              =   6900
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   6810
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   1020
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่ออำเภอ/ จังหวัด"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   1710
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาคลองส่งน้ำ"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   3120
      Width           =   1665
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frAumphurSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8640
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Index           =   1
      Left            =   0
      Top             =   1590
      Width           =   8685
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ผลการค้นหา"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   -120
      TabIndex        =   5
      Top             =   1710
      Width           =   8760
   End
End
Attribute VB_Name = "frmAumphurSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim colSearchResult_ As Collection
Dim firstActive  As Boolean

Public Sub getCriteria(criStr As String)
    firstActive = True
    txtCriteria.Text = Trim(criStr)
End Sub

Private Sub Form_Activate()
    If firstActive Then
        cmdSearch_Click
        firstActive = False
        If lvResult.ListItems.Count = 1 Then
            callerForm_.SetData colSearchResult_(1)
            Unload Me
        End If
    End If
End Sub

Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub cmdOpen_Click()
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีข้อมูลอำเภอที่ต้องการที่ต้องการ", vbExclamation
        Exit Sub
    End If
    If lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกอำเภอที่ต้องการที่ต้องการ", vbExclamation
        Exit Sub
    End If
    callerForm_.SetData colSearchResult_(lvResult.SelectedItem.Key)
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim criStr As String
Dim strCond As String
Dim x%
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As IMIAumphur
    criStr = Replace(Replace(Trim(txtCriteria.Text), "*", "%"), "'", "")
    If Trim(criStr) = vbNullString Then
        MsgBox "กรุณาใส่คำในการค้นหา", vbExclamation
        txtCriteria.SetFocus
        Exit Sub
    End If
    Set colSearchResult_ = Nothing
    Set colSearchResult_ = New Collection
    lvResult.ListItems.Clear
    Set tmpS = New BWGSearchManager
    strCond = "(ProvinceName Like '%" & criStr & "%' Or AumphurName Like '%" & criStr & "%')"
    
    Set colSearchResult_ = tmpS.AumphurSearch(strCond)
    Set tmpS = Nothing
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.ProvinceName)
        tmpList.SubItems(1) = Trim(tmpItem.Name)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
End Sub

Private Sub Form_Load()
    With frmPicture
        cmdSearch.Picture = .picSearch.Picture
        cmdClose.Picture = .picExit.Picture
    End With
End Sub

Public Sub SetCallerForm(frm As Form)
    Set callerForm_ = frm
End Sub

Private Sub lvResult_DblClick()
    cmdOpen_Click
End Sub

Private Sub lvResult_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpItem As IMIAumphur
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    
    If KeyCode = 46 Then
        'If CheckPwdBeforeDelete = False Then Exit Sub
        If MsgBox("ต้องการลบข้อมูลรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            Set tmpItem = New IMIAumphur
            tmpItem.ID = Replace(lvResult.SelectedItem.Key, "'", "")
            tmpItem.Delete
            lvResult.ListItems.Remove lvResult.SelectedItem.Index
            Set tmpItem = Nothing
        End If
    End If
    If KeyCode = 13 Then cmdOpen_Click

End Sub

Private Sub txtCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
End Sub

Private Sub txtCriteria_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub
