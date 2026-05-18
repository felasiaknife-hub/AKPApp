VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTransportRateSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F127"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   45
      Picture         =   "frTransportRateSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5235
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   8685
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5235
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   7740
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5235
      Width           =   915
   End
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1620
      TabIndex        =   0
      Text            =   "*"
      Top             =   900
      Width           =   4305
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3375
      Left            =   0
      TabIndex        =   4
      Top             =   1830
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   5953
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ต้นทาง"
         Object.Width           =   4339
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ปลายทาง"
         Object.Width           =   4022
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ประเภทรถ"
         Object.Width           =   3889
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "จ่าย Vendor"
         Object.Width           =   2487
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "เก็บลูกค้า"
         Object.Width           =   2487
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "อำเภอ-จังหวัด ต้นทางและปลายทาง, ประเภทรถ"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   1740
      TabIndex        =   8
      Top             =   1260
      Width           =   3225
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
      Left            =   -90
      TabIndex        =   7
      Top             =   1605
      Width           =   9765
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Index           =   1
      Left            =   -30
      Top             =   1515
      Width           =   9765
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   6
      Top             =   945
      Width           =   1365
   End
   Begin VB.Label AumphurSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาอัตราค่าขนส่ง"
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
      Left            =   285
      TabIndex        =   5
      Top             =   30
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   195
      X2              =   6000
      Y1              =   615
      Y2              =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   195
      X2              =   5910
      Y1              =   525
      Y2              =   525
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   -15
      Picture         =   "frTransportRateSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9660
   End
End
Attribute VB_Name = "frmTransportRateSearch"
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
            callerForm_.SetDataTransportRate colSearchResult_(1)
            Unload Me
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีข้อมูลพนักงานที่ต้องการ", vbExclamation
        Exit Sub
    End If
    If lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกพนักงานที่ต้องการ", vbExclamation
        Exit Sub
    End If
    callerForm_.SetDataTransportRate colSearchResult_(lvResult.SelectedItem.Key)
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim criStr As String
Dim strCond As String
Dim x%, result
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGTransVendorChargeRate
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
    strCond = "(FromAumphurName Like '%" & criStr & "%' Or FromProvinceName Like '%" & criStr & "%' Or ToAumphurName Like '%" & criStr & "%' Or ToProvinceName Like '%" & criStr & "%' Or TruckTypeDesc Like '%" & criStr & "%')"
    Set colSearchResult_ = tmpS.TransVendorChargeRateSearch(strCond)
    If colSearchResult_ Is Nothing Then Exit Sub
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.FromAumphurName & " - " & tmpItem.FromProvinceName)
        tmpList.SubItems(1) = Trim(tmpItem.ToAumphurName & " - " & tmpItem.ToProvinceName)
        tmpList.SubItems(2) = Trim(tmpItem.CarTypeName & " - " & tmpItem.CarTypeName)
        tmpList.SubItems(3) = Trim(tmpItem.RatePerKM & " - " & tmpItem.RatePerTrip)
        tmpList.SubItems(4) = Trim(tmpItem.CustRatePerKM & " - " & tmpItem.CustRatePerTrip)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        If firstActive = True Then
            result = MsgBox("ไม่พบข้อมูลที่ต้องการค้นหา ต้องการเพิ่มข้อมูลอัตราค่าขนส่งใหม่หรือไม่", vbQuestion + vbYesNo)
            If result = vbYes Then
                frmTransportRate.cmdSearch.Enabled = False
                frmTransportRate.Show 1
                Set colSearchResult_ = tmpS.TransVendorChargeRateSearch(strCond)
                For x = 1 To colSearchResult_.Count
                    Set tmpItem = colSearchResult_(x)
                    Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.FromAumphurName & " - " & tmpItem.FromProvinceName)
                    tmpList.SubItems(1) = Trim(tmpItem.ToAumphurName & " - " & tmpItem.ToProvinceName)
                    tmpList.SubItems(2) = Trim(tmpItem.CarTypeName & " - " & tmpItem.CarTypeName)
                    tmpList.SubItems(3) = Trim(tmpItem.RatePerKM & " - " & tmpItem.RatePerTrip)
                    tmpList.SubItems(4) = Trim(tmpItem.CustRatePerKM & " - " & tmpItem.CustRatePerTrip)
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

Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    MsgBox ColumnHeader.Width
End Sub

Private Sub txtCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
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
    If KeyCode = 13 Then cmdOpen_Click
End Sub
