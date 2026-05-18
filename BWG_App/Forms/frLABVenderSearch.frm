VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmLABVenderSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F055"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1575
      TabIndex        =   0
      Text            =   "*"
      Top             =   990
      Width           =   4305
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   5550
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5505
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6495
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5505
      Width           =   915
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   45
      Picture         =   "frLABVenderSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5505
      Width           =   915
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3465
      Left            =   15
      TabIndex        =   1
      Top             =   1980
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   6112
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
         Text            =   "ชื่อบริษัท"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ที่อยู่"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   90
      TabIndex        =   8
      Top             =   1035
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   30
      X2              =   5745
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   30
      X2              =   5835
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label AumphurSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาบริษัท"
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
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   75
      Width           =   1035
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   30
      Index           =   1
      Left            =   -30
      Top             =   1665
      Width           =   7515
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
      Left            =   -150
      TabIndex        =   6
      Top             =   1725
      Width           =   7635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัสบริษัท, ชื่อบริษัท, ชื่อกลุ่มบริษัท"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   1680
      TabIndex        =   5
      Top             =   1350
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   -15
      Picture         =   "frLABVenderSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "frmLABVenderSearch"
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
            callerForm_.SetDataVenderLAB colSearchResult_(1)
            Unload Me
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีข้อมูลลูกค้าที่ต้องการ", vbExclamation
        Exit Sub
    End If
    If lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกลูกค้าที่ต้องการ", vbExclamation
        Exit Sub
    End If
    callerForm_.SetDataVenderLAB colSearchResult_(lvResult.SelectedItem.Key)
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim criStr As String
Dim strCond As String
Dim x%, result
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGTransportVendor
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
    strCond = "(NameTH Like '%" & criStr & "%' AND VendorType = 'LB')"
    Set colSearchResult_ = tmpS.TransportVenderSearch(strCond)
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.NameTH)
        tmpList.SubItems(1) = Trim(tmpItem.TumbolName & " " & tmpItem.AumphurName & " " & tmpItem.ProvinceName)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        If firstActive = True Then
            result = MsgBox("ไม่พบข้อมูลที่ต้องการค้นหา ต้องการเพิ่มข้อมูลใหม่หรือไม่", vbQuestion + vbYesNo)
            If result = vbYes Then
                frmLABVender.cmdSearch.Enabled = False
                frmLABVender.Show 1
                Set colSearchResult_ = tmpS.TransportVenderSearch(strCond)
                For x = 1 To colSearchResult_.Count
                    Set tmpItem = colSearchResult_(x)
                    Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.NameTH)
                    tmpList.SubItems(1) = Trim(tmpItem.TumbolName & " " & tmpItem.AumphurName & " " & tmpItem.ProvinceName)
                    Set tmpItem = Nothing
                    Set tmpList = Nothing
                Next
            End If
        Else
            MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
        End If
    End If
    Set tmpS = Nothing
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
On Error Resume Next
    cmdOpen_Click
End Sub

Private Sub lvResult_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpItem As BWGTransportVendor
Dim tmpStr As String
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    
    If KeyCode = 46 Then
        If MsgBox("ต้องการลบข้อมูลรายการนี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
            Set tmpItem = New BWGTransportVendor
            tmpItem.ID = Replace(lvResult.SelectedItem.Key, "'", "")
            tmpStr = tmpItem.Delete
            If Trim(tmpStr) = "" Then
                lvResult.ListItems.Remove lvResult.SelectedItem.Index
                Set tmpItem = Nothing
                txtCriteria.Text = ""
                txtCriteria.SetFocus
            Else
                MsgBox tmpStr, vbExclamation
            End If
        End If
    End If
    If KeyCode = 13 Then cmdOpen_Click

End Sub

Private Sub txtCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click

End Sub

