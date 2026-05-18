VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmFSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ค้นหาข้อมูล"
   ClientHeight    =   5670
   ClientLeft      =   7380
   ClientTop       =   2940
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6870
   Begin VB.ComboBox cbQ 
      Height          =   315
      Left            =   1980
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   585
      Width           =   1410
   End
   Begin VB.TextBox TxtYear 
      Height          =   285
      Left            =   3780
      TabIndex        =   8
      Top             =   585
      Width           =   1185
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "เปิดข้อมูล"
      Height          =   675
      Left            =   60
      Picture         =   "FrmFSearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4905
      Width           =   915
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ปิด"
      Height          =   675
      Left            =   5895
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4905
      Width           =   915
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหา"
      Height          =   675
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4905
      Width           =   915
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   3585
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1275
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   6324
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
         Text            =   "ไตรมาสที่"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ปี"
         Object.Width           =   4233
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
      Height          =   285
      Index           =   2
      Left            =   -90
      TabIndex        =   7
      Top             =   1005
      Width           =   6930
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   45
      Index           =   1
      Left            =   0
      Top             =   945
      Width           =   7080
   End
   Begin VB.Label AumphurSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหาข้อมูลเก่า"
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
      TabIndex        =   6
      Top             =   -45
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   30
      X2              =   5835
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   30
      X2              =   5745
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ปี :"
      Height          =   195
      Index           =   7
      Left            =   3615
      TabIndex        =   5
      Top             =   615
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Q ที่"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   1665
      TabIndex        =   4
      Top             =   630
      Width           =   270
   End
   Begin VB.Image Image2 
      Height          =   660
      Left            =   -15
      Picture         =   "FrmFSearch.frx":27A2
      Stretch         =   -1  'True
      Top             =   -135
      Width           =   7080
   End
End
Attribute VB_Name = "FrmFSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim callerForm_ As Form
Dim colSearchResult_ As Collection
Dim firstActive  As Boolean
Dim strOptional As String
Dim ShowSubContact_ As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
                Dim tSql As String
                Dim RS As New ADODB.Recordset
                Dim i As Integer
                    tSql = "SELECT * FROM Financial_Ratios " & vbCrLf
                    tSql = tSql & "where id= " & lvResult.SelectedItem.Key & " "
                    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    With FrmFinancialRatio
        .cbQ.Text = RS!q
        .cbYears.Text = RS!qname
        .RevenueBuring.Text = Format(RS!RevenueBuring, "#,##0.00")
        .Transportation.Text = Format(RS!TransportationRubbish, "#,##0.00")
        .OtherIncome.Text = Format(RS!OtherIncome, "#,##0.00")
        .NetIncome.Text = Format(RS!NetIncome, "#,##0.00")
        .CostOfService.Text = Format(RS!CostOfService, "#,##0.00")
        .GrossProfit.Text = Format(RS!GrossProfit, "#,##0.00")
        .Expense.Text = Format(RS!Expenses, "#,##0.00")
        .EBITDA.Text = Format(RS!EBITDA, "#,##0.00")
        .EBIT.Text = Format(RS!EBIT, "#,##0.00")
        .ShareofProfit.Text = Format(RS!ShareofProfit, "#,##0.00")
        .NetProfit.Text = Format(RS!NetProfit, "#,##0.00")
        .ROE.Text = Format(RS!ROE, "#,##0.0000")
        .ROA.Text = Format(RS!ROA, "#,##0.0000")
        .DE.Text = Format(RS!dex, "#,##0.0000")
        .DebtRatio.Text = Format(RS!DebtRatio, "#,##0.0000")
        .BV.Text = Format(RS!BV, "#,##0.0000")
        .EPS.Text = Format(RS!EPS, "#,##0.0000")
        .Asset.Text = Format(RS!Assete, "#,##0.00")
        .Liability.Text = Format(RS!Liabilities, "#,##0.00")
        .Equity.Text = Format(RS!Equity, "#,##0.00")
        .OperatingAct.Text = Format(RS!OperatingAc, "#,##0.00")
        .InvestingAct.Text = Format(RS!InvestingAc, "#,##0.00")
        .FinancingAct.Text = Format(RS!FinancingAc, "#,##0.00")
        .NetIncrease.Text = Format(RS!NetIncrease, "#,##0.00")
        .EndingCash.Text = Format(RS!EndingCash, "#,##0.00")
        .lblNew.Caption = "N"
        .lblID.Caption = Replace(lvResult.SelectedItem.Key, "'", "")
    End With
        Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
Dim criStr As String
Dim strCond As String
Dim x%, result
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim tmpItem As BWGCustomer
    Set colSearchResult_ = Nothing
    Set colSearchResult_ = New Collection
    lvResult.ListItems.Clear
                Dim tSql As String
                Dim RS As New ADODB.Recordset
                Dim i As Integer
                    tSql = "SELECT * FROM Financial_Ratios " & vbCrLf
                    If cbQ.ListIndex > 0 Then
                    tSql = tSql & "where q= '" & Trim(cbQ.Text) & "' " & vbCrLf
                        If Trim(TxtYear.Text) <> "" Then
                            tSql = tSql & " and qname='" & Trim(TxtYear.Text) & "' "
                        End If
                    ElseIf Trim(TxtYear.Text) <> "" Then
                            tSql = tSql & " where qname='" & Trim(TxtYear.Text) & "' "
                    End If
                    tSql = tSql & "order by qname,q"
                    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                        For x = 1 To RS.RecordCount
                            Set tmpList = lvResult.ListItems.Add(, "'" & RS!ID & "'", RS!q)
                            tmpList.SubItems(1) = Trim("" & RS!qname)
                            Set tmpList = Nothing
                            RS.MoveNext
                        Next
    If lvResult.ListItems.Count = 0 Then
            MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
    Set tmpS = Nothing

End Sub

Private Sub Form_Load()
    Dim i As Integer
    cbQ.AddItem "เลือก"
    For i = 1 To 4
        cbQ.AddItem i
    Next
    cbQ.ListIndex = 0
End Sub

Private Sub lvResult_DblClick()
    cmdOpen_Click
End Sub
