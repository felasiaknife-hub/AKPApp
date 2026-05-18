VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCarDaily 
   Caption         =   "รายการจองรถในแต่ละเดือน"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8535
      Left            =   0
      ScaleHeight     =   8505
      ScaleWidth      =   2625
      TabIndex        =   1
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton CmdExp 
         Caption         =   "Excel"
         Height          =   435
         Left            =   135
         TabIndex        =   6
         Top             =   2295
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   135
         TabIndex        =   5
         Top             =   2790
         Width           =   2295
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   135
         TabIndex        =   4
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox CBMonth 
         Height          =   315
         ItemData        =   "FrmCarDaily.frx":0000
         Left            =   225
         List            =   "FrmCarDaily.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   405
         Width           =   1410
      End
      Begin VB.ComboBox CbYear 
         Height          =   315
         ItemData        =   "FrmCarDaily.frx":009A
         Left            =   225
         List            =   "FrmCarDaily.frx":009C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   990
         Width           =   1410
      End
      Begin VB.Line Line2 
         X1              =   135
         X2              =   2445
         Y1              =   1710
         Y2              =   1710
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เดือน"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   8
         Top             =   150
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปี"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   765
         Width           =   105
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   32
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ประเภทรถ"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "1"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "2"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "3"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "4"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "5"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "6"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "7"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "8"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "9"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "10"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "11"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "12"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "13"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "14"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "15"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "16"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "17"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "18"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "19"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "20"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "21"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "22"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "23"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "24"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "25"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "26"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "27"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "28"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "29"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "30"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "31"
         Object.Width           =   882
      EndProperty
   End
End
Attribute VB_Name = "FrmCarDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaleStaff_ As BWGCompanyStaff
Dim curCustomer_ As BWGCustomer
Dim colSearchResult_ As Collection

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdExp_Click()
'    Dim objExcel As New Excel.Application
'
'    Dim objExcelSheet As Excel.Worksheet
'    Dim col, row As Integer
'
'    If lvResult.ListItems.Count > 0 Then
'        objExcel.Workbooks.Add
'        Set objExcelSheet = objExcel.Worksheets.Add
'
'
'        For col = 1 To lvResult.ColumnHeaders.Count
'            objExcelSheet.Cells(1, col).Value = lvResult.ColumnHeaders(col)
'        Next
'
'        For row = 2 To lvResult.ListItems.Count + 1
'           For col = 1 To lvResult.ColumnHeaders.Count
'           If col = 1 Then
'                   objExcelSheet.Cells(row, col).Value = lvResult.ListItems(row - 1).Text
'           Else
'                   objExcelSheet.Cells(row, col).Value = lvResult.ListItems(row - 1).SubItems(col - 1)
'           End If
'           Next
'        Next row
'
'        objExcelSheet.Columns.AutoFit
'
'
'        objExcel.Visible = True
'
'    Else
'        MsgBox "No data to export", vbInformation, Me.Caption
'    End If

End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrD
    Dim criStr As String
    Dim strCond As String
    Dim tmpRec As New ADODB.Recordset
    Dim x%, result%, Y%, i%
    Dim tmpList As ListItem
    Dim tmpS As BWGSearchManager
    Dim tmpItem As BWGQuotationHeader
    Dim tmpQuoDet As BWGQuotationDetail
    Dim tmpColDet As Collection
    
    Screen.MousePointer = 0
    Set colSearchResult_ = Nothing
    Set colSearchResult_ = New Collection
    lvResult.ListItems.Clear
    Set tmpS = New BWGSearchManager
    Dim mm As String
    mm = CBMonth.ListIndex + 1
    If Len(mm) = 1 Then
        mm = "0" & mm
    Else
        mm = mm
    End If
        sqlStr = "select TruckTypeDesc, "
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/01" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D1,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/02" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D2,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/03" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D3,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/04" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D4,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/05" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D5,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/06" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D6,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/07" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D7,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/08" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D8,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/09" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D9,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/10" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D10,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/11" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D11,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/12" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D12,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/13" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D13,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/14" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D14,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/15" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D15,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/16" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D16,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/17" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D17,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/18" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D18,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/19" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D19,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/20" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D20,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/21" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D21,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/22" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D22,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/23" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D23,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/24" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D24,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/25" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D25,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/26" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D26,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/27" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D27,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/28" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D28,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/29" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D29,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/30" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D30,"
        sqlStr = sqlStr & "(select sum(TruckCount) as Tcountfrom from vw_JobListForCarAssigned jd where WorkDate='" & Trim(CbYear.Text) & "/" & mm & "/31" & "' and jd.TruckTypeDesc=vw_JobListForCarAssigned.TruckTypeDesc  group by TruckTypeDesc,WorkDate) as D31"
        sqlStr = sqlStr & " from vw_JobListForCarAssigned where WorkDate>='" & Trim(CbYear.Text) & "/" & mm & "/01" & "' and WorkDate<='" & Trim(CbYear.Text) & "/" & mm & "/31" & "' group by TruckTypeDesc order by TruckTypeDesc"
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    
    If tmpRec.RecordCount = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา กรุณาตรวจสอบ", vbOKOnly + vbInformation, "ไม่มีข้อมูล"
        Screen.MousePointer = 0
        Exit Sub
    End If
    DoEvents
    Screen.MousePointer = 11
    Do While Not tmpRec.EOF
        Dim ForeColor
        Set tmpList = lvResult.ListItems.Add(, , tmpRec!TruckTypeDesc)
        If tmpRec!d1 <> "Null" Then tmpList.SubItems(1) = tmpRec!d1
        If tmpRec!d2 <> "Null" Then tmpList.SubItems(2) = tmpRec!d2
        If tmpRec!d3 <> "Null" Then tmpList.SubItems(3) = tmpRec!d3
        If tmpRec!d4 <> "Null" Then tmpList.SubItems(4) = tmpRec!d4
        If tmpRec!d5 <> "Null" Then tmpList.SubItems(5) = tmpRec!d5
        If tmpRec!d6 <> "Null" Then tmpList.SubItems(6) = tmpRec!d6
        If tmpRec!d7 <> "Null" Then tmpList.SubItems(7) = tmpRec!d7
        If tmpRec!d8 <> "Null" Then tmpList.SubItems(8) = tmpRec!d8
        If tmpRec!d9 <> "Null" Then tmpList.SubItems(9) = tmpRec!d9
        If tmpRec!d10 <> "Null" Then tmpList.SubItems(10) = tmpRec!d10
        If tmpRec!d11 <> "Null" Then tmpList.SubItems(11) = tmpRec!d11
        If tmpRec!d12 <> "Null" Then tmpList.SubItems(12) = tmpRec!d12
        If tmpRec!d13 <> "Null" Then tmpList.SubItems(13) = tmpRec!d13
        If tmpRec!d14 <> "Null" Then tmpList.SubItems(14) = tmpRec!d14
        If tmpRec!d15 <> "Null" Then tmpList.SubItems(15) = tmpRec!d15
        If tmpRec!d16 <> "Null" Then tmpList.SubItems(16) = tmpRec!d16
        If tmpRec!d17 <> "Null" Then tmpList.SubItems(17) = tmpRec!d17
        If tmpRec!d18 <> "Null" Then tmpList.SubItems(18) = tmpRec!d18
        If tmpRec!d19 <> "Null" Then tmpList.SubItems(19) = tmpRec!d19
        If tmpRec!d20 <> "Null" Then tmpList.SubItems(20) = tmpRec!d20
        If tmpRec!d21 <> "Null" Then tmpList.SubItems(21) = tmpRec!d21
        If tmpRec!d22 <> "Null" Then tmpList.SubItems(22) = tmpRec!d22
        If tmpRec!d23 <> "Null" Then tmpList.SubItems(23) = tmpRec!d23
        If tmpRec!d24 <> "Null" Then tmpList.SubItems(24) = tmpRec!d24
        If tmpRec!d25 <> "Null" Then tmpList.SubItems(25) = tmpRec!d25
        If tmpRec!d26 <> "Null" Then tmpList.SubItems(26) = tmpRec!d26
        If tmpRec!d27 <> "Null" Then tmpList.SubItems(27) = tmpRec!d27
        If tmpRec!d28 <> "Null" Then tmpList.SubItems(28) = tmpRec!d28
        If tmpRec!d29 <> "Null" Then tmpList.SubItems(29) = tmpRec!d29
        If tmpRec!d30 <> "Null" Then tmpList.SubItems(30) = tmpRec!d30
        If tmpRec!d31 <> "Null" Then tmpList.SubItems(31) = tmpRec!d31
        DoEvents
        Set tmpList = Nothing
        tmpRec.MoveNext
    Loop
    Set tmpRec = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
ErrD:
    MsgBox err.Description, vbCritical
    Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
    Dim Y As String
    Dim AllYear As String
    Dim i As Integer
    If Year(Date) < 2550 Then
        Y = Year(Date) + 543
    Else
        Y = Year(Date)
    End If
    For i = 0 To 3
        AllYear = (Y - 2) + i
        CbYear.AddItem AllYear
    Next
    CbYear.ListIndex = 2
    Dim M As Integer
    M = Month(Date)
        CBMonth.ListIndex = M - 1
End Sub

Private Sub Form_Resize()
On Error Resume Next
    With lvResult
        If Me.Width - Picture1.Width > 500 Then
            .Width = Me.Width - Picture1.Width - 200
        End If
        .Height = Me.Height - .Top - 500
    End With
End Sub
