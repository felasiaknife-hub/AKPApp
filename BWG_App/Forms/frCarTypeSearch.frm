VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCarTypeSearch 
   BackColor       =   &H00FFFFFF&
   Caption         =   "F005"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCarType 
      Height          =   315
      Left            =   5370
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   900
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   795
      Left            =   45
      Picture         =   "frCarTypeSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5235
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   795
      Left            =   6045
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5235
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   795
      Left            =   5100
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
      Width           =   3645
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3375
      Left            =   0
      TabIndex        =   4
      Top             =   1830
      Width           =   7005
      _ExtentX        =   12356
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ประเภทกลุ่มหลัก"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ประเภทกลุ่มย่อย"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ประเภทกลุ่มหลัก, ประเภทกลุ่มย่อย"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   1740
      TabIndex        =   8
      Top             =   1260
      Width           =   2400
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
      Width           =   7125
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   60
      Index           =   1
      Left            =   -30
      Top             =   1515
      Width           =   7125
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
      Caption         =   "ค้นหาประเภทรถ"
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
      Left            =   75
      TabIndex        =   5
      Top             =   0
      Width           =   1635
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   -15
      X2              =   5790
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   -15
      X2              =   5700
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   -15
      Picture         =   "frCarTypeSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7020
   End
End
Attribute VB_Name = "frmCarTypeSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim colSearchResult_ As Collection
Dim selCarTypeID_ As String
Dim vendorID_ As String
Dim isTrailer_ As Boolean
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
            callerForm_.SetDataTruckSubType colSearchResult_(1)
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
    callerForm_.SetDataTruckSubType colSearchResult_(lvResult.SelectedItem.Key), isTrailer_
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim criStr As String
Dim strCond As String
Dim x%, result%
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGTruckSubType
    criStr = Replace(Replace(Trim(txtCriteria.Text), "*", "%"), "'", "")
    If Trim(criStr) = vbNullString Then
        MsgBox "กรุณาใส่คำในการค้นหา", vbExclamation
        txtCriteria.SetFocus
        Exit Sub
    End If
    If cboCarType.ListIndex = 1 Then
        isTrailer_ = True
    Else
        isTrailer_ = False
    End If
    Set colSearchResult_ = Nothing
    Set colSearchResult_ = New Collection
    lvResult.ListItems.Clear
    Set tmpS = New BWGSearchManager
    strCond = "(SubTypeDesc Like '%" & criStr & "%' Or TruckTypeDesc Like '%" & criStr & "%')"
    If Trim(selCarTypeID_) <> "" Then
        strCond = strCond & " And TruckTypeID='" & selCarTypeID_ & "'"
    End If
    If Trim(vendorID_) <> "" Then
        strCond = strCond & " And SubTypeID in (Select SubTypeID from tbVendorCar WHERE VendorID='" & vendorID_ & "')"
    End If
'    If isTrailer_ Then
'        strCond = strCond & " And CanBeTrailer='Y'"
'    End If
    Set colSearchResult_ = tmpS.TruckSubTypeSearch(strCond)
    Set tmpS = Nothing
    For x = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(x)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.id & "'", tmpItem.TypeDesc)
        If isTrailer_ Then
            tmpList.Text = tmpList.Text & " (พ่วง)"
        End If
        tmpList.SubItems(1) = Trim(tmpItem.SubTypeDesc)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
    Set tmpS = Nothing

End Sub

Private Sub txtCriteria_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
End Sub

Private Sub Form_Load()
    With frmPicture
        cmdSearch.Picture = .picSearch.Picture
        cmdClose.Picture = .picExit.Picture
    End With
    With cboCarType
        .AddItem "รถเดี่ยว"
        .AddItem "รถพ่วง"
        .ListIndex = 0
    End With
End Sub

Public Sub SetCallerForm(frm As Form, Optional selCarTypeID As String, Optional VendorID As String, Optional isTrailer As Boolean, Optional CanBeTrailer As Boolean)
    Set callerForm_ = frm
    selCarTypeID_ = selCarTypeID
    vendorID_ = VendorID
    isTrailer_ = isTrailer
    If isTrailer_ Then
        cboCarType.ListIndex = 1
    Else
        cboCarType.ListIndex = 0
    End If
    'cboCarType.Visible = canBeTrailer_
End Sub

Private Sub lvResult_DblClick()
    cmdOpen_Click
End Sub

Private Sub lvResult_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdOpen_Click
End Sub
