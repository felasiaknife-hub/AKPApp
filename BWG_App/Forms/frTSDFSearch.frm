VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTSDFSearch 
   Caption         =   "F136"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6405
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   5370
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5580
      Width           =   915
   End
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1710
      TabIndex        =   6
      Text            =   "*"
      Top             =   900
      Width           =   4305
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6330
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5580
      Width           =   915
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   60
      Picture         =   "frTSDFSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5580
      Width           =   915
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3645
      Left            =   30
      TabIndex        =   1
      Top             =   1890
      Width           =   7245
      _ExtentX        =   12779
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ชื่อบริษัท"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขประจำตัว"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ที่อยุ่"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อบริษัท, เลขประจำตัว, ที่อยู่"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   1740
      TabIndex        =   8
      Top             =   1260
      Width           =   1980
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   150
      TabIndex        =   7
      Top             =   960
      Width           =   1365
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
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   90
      X2              =   6900
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label AumphurSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาบริษัท TSDF"
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
      Left            =   135
      TabIndex        =   4
      Top             =   0
      Width           =   1875
   End
   Begin VB.Image Image2 
      Height          =   705
      Left            =   0
      Picture         =   "frTSDFSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7320
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
      Left            =   -30
      TabIndex        =   3
      Top             =   1650
      Width           =   7365
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   1
      Left            =   -30
      Top             =   1530
      Width           =   7365
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
      Left            =   660
      TabIndex        =   2
      Top             =   3060
      Width           =   1695
   End
End
Attribute VB_Name = "frmTSDFSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim firstActive As Boolean
Dim callerForm_ As Form
Dim colSearchResult_ As Collection

Public Sub SetCallerForm(frm As Form)
    Set callerForm_ = frm
End Sub

Public Sub getCriteria(criStr As String)
    firstActive = True
    txtCriteria.Text = Trim(criStr)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    If lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกข้อมูลที่ต้องการ", vbExclamation
        Exit Sub
    End If
    callerForm_.SetDataTSDF colSearchResult_(lvResult.SelectedItem.Key)
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    cmdSearchTumbol
End Sub

Private Sub Form_Activate()
    If firstActive Then
        cmdSearchTumbol
        firstActive = False
        If lvResult.ListItems.Count = 1 Then
            callerForm_.SetDataTSDF colSearchResult_(1)
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    With frmPicture
        cmdSearch.Picture = .picSearch.Picture
        cmdClose.Picture = .picExit.Picture
    End With
End Sub

Private Sub cmdSearchTumbol()
Dim criStr As String
Dim result As Integer
Dim strCond As String
Dim x%
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGTSDF
    criStr = Replace(Replace(Trim(txtCriteria.Text), "*", "%"), "'", "")
    If Trim(criStr) = vbNullString Then
        MsgBox "กรุณาใส่คำในการค้นหา", vbExclamation
        Unload Me
        Exit Sub
    End If
    Set colSearchResult_ = Nothing
    Set colSearchResult_ = New Collection
    lvResult.ListItems.Clear
    Set tmpS = New BWGSearchManager
    strCond = "(CompanyName Like '%" & criStr & "%' Or TSDFRegisNo Like '%" & criStr & "%' Or TumbolName Like '%" & criStr & "%' Or AumphurName Like '%" & criStr & "%' Or ProvinceName Like '%" & criStr & "%')"
    
    Set colSearchResult_ = tmpS.TSDFSearch(strCond)
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.CompanyName)
        tmpList.SubItems(1) = Trim(tmpItem.TSDFRegisNo)
        tmpList.SubItems(2) = Trim(tmpItem.Addr1 & " " & tmpItem.TumbolName & " " & tmpItem.AumphurName & " " & tmpItem.ProvinceName)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbInformation
'        result = MsgBox("ไม่พบข้อมูลที่ต้องการค้นหา ต้องการเพิ่มข้อมูลบริษัท TSDF ใหม่หรือไม่", vbQuestion + vbYesNo)
'        If result = vbYes Then
'            frmProvince.Show 1
'            Set colSearchResult_ = tmpS.TSDFSearch(strCond)
'            For x = 1 To colSearchResult_.Count
'                Set tmpItem = colSearchResult_(x)
'                Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.CompanyName)
'                tmpList.SubItems(1) = Trim(tmpItem.TSDFRegisNo)
'                tmpList.SubItems(2) = Trim(tmpItem.Addr1 & " " & tmpItem.TumbolName & " " & tmpItem.AumphurName & " " & tmpItem.ProvinceName)
'                Set tmpItem = Nothing
'                Set tmpList = Nothing
'            Next
'        End If
    End If
    Set tmpS = Nothing
End Sub

Private Sub lvResult_DblClick()
On Error Resume Next
    cmdOpen_Click
End Sub

Private Sub lvResult_KeyDown(KeyCode As Integer, Shift As Integer)
Dim tmpItem As IMITumbol
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    If KeyCode = 13 Then cmdOpen_Click
End Sub

Private Sub txtCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then cmdSearch_Click
End Sub
