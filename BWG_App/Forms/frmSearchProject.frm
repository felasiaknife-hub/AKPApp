VERSION 5.00
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearchProject 
   Caption         =   "ค้นหาข้อมูลโปรเจค"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   10155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "ปิด"
      Height          =   735
      Left            =   9000
      TabIndex        =   9
      Top             =   5850
      Width           =   1140
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   735
      Left            =   7065
      TabIndex        =   2
      Top             =   180
      Width           =   2130
   End
   Begin VB.TextBox txtSearch 
      Height          =   330
      Left            =   2295
      TabIndex        =   1
      Top             =   180
      Width           =   4605
   End
   Begin DateCtl.Date dtFrom 
      Height          =   315
      Left            =   2430
      TabIndex        =   3
      Top             =   585
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin DateCtl.Date dtTo 
      Height          =   315
      Left            =   4950
      TabIndex        =   4
      Top             =   585
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      BackColor       =   -2147483633
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   4455
      Left            =   0
      TabIndex        =   7
      Top             =   1335
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   7858
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
         Text            =   "ชื่อโปรเจค"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่เริ่ม"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่เสร็จ"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "รหัสลูกค้า"
         Object.Width           =   1677
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   8820
      EndProperty
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
      TabIndex        =   8
      Top             =   1095
      Width           =   10320
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   1
      Left            =   0
      Top             =   1035
      Width           =   10200
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ในช่วงวันที่เริ่ม :"
      Height          =   195
      Index           =   0
      Left            =   1125
      TabIndex        =   6
      Top             =   645
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ถึงวันที่"
      Height          =   195
      Index           =   1
      Left            =   4350
      TabIndex        =   5
      Top             =   645
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "รหัส หรือ ชื่อลูกค้า : "
      Height          =   195
      Index           =   0
      Left            =   855
      TabIndex        =   0
      Top             =   270
      Width           =   1365
   End
End
Attribute VB_Name = "frmSearchProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Private Sub cmdSearch_Click()
On Error Resume Next
Dim criStr As String
Dim strCond As String
Dim x%, result
Dim tmpList As ListItem
    lvResult.ListItems.Clear
                Dim tSql As String
                Dim rs As New ADODB.Recordset
                Dim i As Integer
                    tSql = "SELECT * FROM vw_JobProject " & vbCrLf
                    tSql = tSql & "where StartDate>= '" & dtFrom.ValueYMD & "' AND StartDate<='" & dtTo.ValueYMD & "' " & vbCrLf
                    If txtSearch.Text <> "" Then
                    tSql = tSql & " and CustomerCode like '%" & Trim(txtSearch.Text) & "%' or CompanyName like '%" & Trim(txtSearch.Text) & "%' " & vbCrLf
                    End If
                    tSql = tSql & "order by StartDate"
                    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                        For x = 1 To rs.RecordCount
                            Set tmpList = lvResult.ListItems.Add(, "'" & rs!ProjectID & "'", rs!ProjectName)
                            tmpList.SubItems(1) = Trim("" & rs!StartDate)
                            tmpList.SubItems(2) = Trim("" & rs!EndDate)
                            tmpList.SubItems(3) = Trim("" & rs!CustomerCode)
                            tmpList.SubItems(4) = Trim("" & rs!CompanyName)
                            Set tmpList = Nothing
                            rs.MoveNext
                        Next
    If lvResult.ListItems.Count = 0 Then
            MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub
Public Sub SetCallerForm(frm As Form)
    Set callerForm_ = frm
End Sub

Private Sub lvResult_DblClick()
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีข้อมูลที่ต้องการ", vbExclamation
        Exit Sub
    End If
    If lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกข้อมูลที่ต้องการ", vbExclamation
        Exit Sub
    End If
    callerForm_.SetdataProject (lvResult.SelectedItem.Key)
'    Unload Me
    Me.Hide

End Sub
