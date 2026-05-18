VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frListMenifest 
   Caption         =   "รายการ Menifest"
   ClientHeight    =   9000
   ClientLeft      =   1620
   ClientTop       =   60
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   11265
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      Begin Commission.ctlDate dtFrom 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtCustomer 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Text            =   "*"
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3960
         Width           =   2235
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   120
         TabIndex        =   1
         Top             =   3240
         Width           =   2235
      End
      Begin Crystal.CrystalReport rpt1 
         Left            =   120
         Top             =   7380
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin Commission.ctlDate dtTo 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         Caption         =   "วันที่ :"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึง"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า , ผู้แทนขาย"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   1560
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3495
      Left            =   2640
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   6165
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ลำดับ"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ลำดับ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "เลข Menifest"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "วันที่"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "เลข INV"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "วันที่ บันทึก INV"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "เลข RE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "วันที่ ออก RE"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "พนังงานขาย"
         Object.Width           =   2822
      EndProperty
   End
End
Attribute VB_Name = "frListMenifest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim colSearchResult_ As Collection
Dim firstActive  As Boolean

Private Sub cmdClose_Click()
     Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Dim criStr As String
Dim strCond As String
Dim X%, Result
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGMenifestBillRE
    criStr = Replace(Replace(Trim(txtCustomer.Text), "*", "%"), "'", "")
    If Trim(criStr) = vbNullString Then
        MsgBox "กรุณาใส่คำในการค้นหา", vbExclamation
        txtCustomer.SetFocus
        Exit Sub
    End If
    Set colSearchResult_ = Nothing
    Set colSearchResult_ = New Collection
    lvResult.ListItems.Clear
    Set tmpS = New BWGSearchManager
    If txtCustomer.Text = Trim("") Or txtCustomer.Text = Trim("*") Then
        strCond = " WorkDate >= '" & dtFrom.ValueYMD & "' And WorkDate <= '" & dtTo.ValueYMD & "'"
    Else
        strCond = "(CustomerName Like '%" & criStr & "%' Or SaleFName Like '%" & criStr & "%') And WorkDate >= '" & dtFrom.ValueYMD & "' And WorkDate <= '" & dtTo.ValueYMD & "'"
    End If
    Set colSearchResult_ = tmpS.MenifestSearchRE(strCond)
    For X = 1 To colSearchResult_.Count
        Set tmpItem = colSearchResult_(X)
         Set tmpList = lvResult.ListItems.Add(, "'" & tmpItem.ID & "'", tmpItem.DocNo)
            tmpList.SubItems(1) = X
            tmpList.SubItems(2) = Trim(tmpItem.DocNo)
            tmpList.SubItems(3) = Trim(tmpItem.WorkDate)
            tmpList.SubItems(4) = Trim(tmpItem.BillingNo)
            tmpList.SubItems(5) = Trim(tmpItem.IssuedDate)
            tmpList.SubItems(6) = Trim(tmpItem.RENo)
            tmpList.SubItems(7) = Trim(tmpItem.REDate)
            tmpList.SubItems(8) = Trim(tmpItem.SaleFName) & " " & Trim(tmpItem.SaleLName)
        Set tmpItem = Nothing
        Set tmpList = Nothing
    Next
'    If lvResult.ListItems.Count = 0 Then
'        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbInformation, "Result"
'    End If
    Set tmpS = Nothing
End Sub

Private Sub Form_Resize()
On Error Resume Next
    lvResult.Width = Me.Width - Frame1.Width - 100
    lvResult.Height = Me.Height - 520
    Frame1.Height = Me.Height - 410
    
'    If lvResult.Width > (lvResult.ColumnHeaders(2).Width + lvResult.ColumnHeaders(3).Width + lvResult.ColumnHeaders(4).Width + 2500) Then
'        lvResult.ColumnHeaders(1).Width = lvResult.Width - (lvResult.ColumnHeaders(2).Width + lvResult.ColumnHeaders(3).Width + lvResult.ColumnHeaders(4).Width) - 500
'    End If
End Sub
