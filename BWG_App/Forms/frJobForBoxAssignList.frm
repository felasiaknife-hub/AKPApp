VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJobForBoxAssignList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "คิวงานรอจัด Box"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   ForeColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   9720
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   4260
      Top             =   4230
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9420
      Left            =   0
      ScaleHeight     =   9390
      ScaleWidth      =   2355
      TabIndex        =   13
      Top             =   0
      Width           =   2385
      Begin VB.CommandButton cmdPrint 
         Caption         =   "พิมพ์รายงาน"
         Height          =   435
         Left            =   60
         TabIndex        =   20
         Top             =   4500
         Width           =   2235
      End
      Begin DateCtl.Date dtTo 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   2760
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin DateCtl.Date dtFrom 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   2100
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtJobNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   8
         Top             =   3510
         Width           =   1995
      End
      Begin VB.ComboBox cboCarSubType 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   2025
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   2025
      End
      Begin VB.CommandButton cmdViewJobBill 
         Caption         =   "ดูข้อมูลใบ Manifest"
         Height          =   435
         Left            =   60
         TabIndex        =   10
         Top             =   8100
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   60
         TabIndex        =   11
         Top             =   4980
         Width           =   2235
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   60
         TabIndex        =   9
         Top             =   4020
         Width           =   2235
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         ItemData        =   "frJobForBoxAssignList.frx":0000
         Left            =   180
         List            =   "frJobForBoxAssignList.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2025
      End
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label lbApprove 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Approve"
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   6180
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   3
         Left            =   120
         Top             =   6540
         Width           =   195
      End
      Begin VB.Label LbBeIncomplete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "be Incomplete"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   360
         TabIndex        =   21
         Top             =   6540
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบสั่งงาน :"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   3270
         Width           =   1080
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   420
         TabIndex        =   7
         Top             =   6900
         Width           =   495
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closed"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   390
         TabIndex        =   19
         Top             =   7260
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   390
         TabIndex        =   18
         Top             =   5880
         Width           =   390
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   120
         Top             =   6900
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   120
         Top             =   7260
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   120
         Top             =   6180
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   120
         Top             =   5850
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   17
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ขนตั้งแต่"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   16
         Top             =   1860
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า, จังหวัดที่ขน :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   14
         Top             =   120
         Width           =   1530
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2400
      TabIndex        =   12
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImgSorted"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบสั่งงาน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "เลขที่ใบคุม"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่ขน"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ทะเบียนรถ"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ลูกค้า"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ผู้ขน"
         Object.Width           =   1162
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ปลายทาง"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "รถที่ใช้ขน"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "สถานที่ขน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "เลขบ็อก"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Box"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "จำนวน"
         Object.Width           =   884
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "กำหนด"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "JobDataCarID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "VendorTruckID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "รับBox"
         Object.Width           =   1764
      EndProperty
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComctlLib.ImageList ImgSorted 
      Left            =   7320
      Top             =   4140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frJobForBoxAssignList.frx":0004
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frJobForBoxAssignList.frx":059E
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อรุ่น,ลาย หรือยี่ห้อ :"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   15
      Top             =   4950
      Width           =   1440
   End
End
Attribute VB_Name = "frmJobForBoxAssignList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCondition_ As String
Dim colCarSubType_ As Collection

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const WS_THICKFRAME         As Long = &H40000
Private Const WS_MAXIMIZE           As Long = &H1000000
Private Const WS_MAXIMIZEBOX        As Long = &H10000
Private Const WS_MINIMIZE           As Long = &H20000000
Private Const WS_MINIMIZEBOX        As Long = &H20000

Private Const WS_EX_WINDOWEDGE      As Long = &H100
Private Const WS_EX_APPWINDOW       As Long = &H40000
Private Const WS_EX_DLGMODALFRAME   As Long = &H1

Private Const GWL_EXSTYLE           As Long = (-20)
Private Const GWL_STYLE             As Long = (-16)

Public Sub RemoveMDIIcon(ByVal lHwnd As Long)
    SetWindowLong lHwnd, GWL_STYLE, GetWindowLong(lHwnd, GWL_STYLE) Or WS_THICKFRAME 'Or WS_MAXIMIZE Or WS_MAXIMIZEBOX 'Or WS_MINIMIZE Or WS_MINIMIZEBOX
    SetWindowLong lHwnd, GWL_EXSTYLE, WS_EX_WINDOWEDGE Or WS_EX_APPWINDOW Or WS_EX_DLGMODALFRAME
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrD
    Dim tmpRec As ADODB.Recordset
    Dim tmpList As ListItem
    Dim x As Long, Y As Long
    Dim sqlStr As String, strCond As String, StatusDesc As String
    Dim tmpArr, CheckAssign As String
    Dim curColor, tripCount As Long
    Dim subStr As String
    
    sqlStr = "SELECT JobID, JobDataCarID, VendorTruckID, JobNo, TimeTableNo, WorkDate, CompanyCode, CompanyName, CASE WHEN WasteType IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
    sqlStr = sqlStr & "THEN 'AKP' ELSE CompanyCode END AS DestCompanyCode, CarRegisID, DriverName, TruckTypeDesc, SubTypeDesc," & vbCrLf
    sqlStr = sqlStr & " AumphurName, ProvinceName, WorkerPerCar , isTrailer, BoxType, BoxCount, AssignedBox" & vbCrLf
    sqlStr = sqlStr & "From dbo.vw_JobListForBoxAssigned" & vbCrLf
    If Trim(txtJobNo.Text) <> "" Then
        strCond = " WHERE JobNo Like '%" & Trim(txtJobNo.Text) & "' "
    Else
        strCond = " WHERE (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "') "
        strCond = strCond & " And (CompanyName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or ProvinceName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')" & vbCrLf
        Select Case cboStatus.ListIndex
            Case 1 'Open
                subStr = " (BoxCount>AssignedBox) "
            Case 2 'Approve
                subStr = " (BoxCount=AssignedBox)  AND ISNULL(BoxCount,0) > 0"
        End Select
        If cboStatus.ListIndex > 0 Then
            If InStr(1, UCase(strCond), "WHERE") > 0 Then
                strCond = strCond & " And " & subStr
            Else
                strCond = strCond & " WHERE " & subStr
            End If
        End If
    
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    strCond = strCond & " AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "') " & vbCrLf
                    strCond = strCond & " AND JobNo NOT LIKE '%PST%'" & vbCrLf
                ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                    strCond = strCond & " AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "') " & vbCrLf
                Else
                    strCond = strCond & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'"
                    'strCond = strCond & " AND WasteType NOT IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
                    strCond = strCond & " AND JobNo NOT LIKE '%PST%'" & vbCrLf
                End If
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    strCond = strCond & " AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "') " & vbCrLf
                    strCond = strCond & " AND JobNo NOT LIKE '%PST%'" & vbCrLf
                ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                    strCond = strCond & " AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "') " & vbCrLf
                Else
                    strCond = strCond & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'"
                    'strCond = strCond & "AND WasteType NOT IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
                    strCond = strCond & "AND JobNo NOT LIKE '%PST%'" & vbCrLf
                End If
            End If
        End If
        
        If cboCarSubType.ListIndex > 0 Then
            strCond = strCond & " And SubTypeID='" & colCarSubType_(cboCarSubType.ListIndex).ID & "'" & vbCrLf
        End If
    End If
    
    strCond = strCond & "AND ISNULL(BoxType, '') <> ''"
    curCondition_ = strCond
    Screen.MousePointer = 11
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\BWTListBox.rpt"
        .SQLQuery = sqlStr & strCond & " ORDER BY JobNo, TimeTableNo, CarRegisID"
        .Destination = crptToWindow
        .PrintReport
        Screen.MousePointer = 0
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub cmdSearch_Click()
    SearchQue
    If lvResult.ListItems.Count = 0 Then MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    If KeyCode = 13 Then
        If Me.ActiveControl.Container.Name = txtKey.Container.Name Then
            If Not TypeOf Me.ActiveControl Is CommandButton Then
                cmdSearch_Click
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim tmpRec As New ADODB.Recordset
    Dim backDate As String
    Dim tmpY As Integer, tmpM As Integer, tmpD As Integer
    Dim tmpArr
    PopulateCompany cboCompany, cboIndex
    PopulateCarSubType
    If IsDate(backDate) Then
        tmpArr = Split(backDate, "/")
        If CInt(tmpArr(0)) < 2500 Then
            dtFrom.ValueYMD = Trim("" & (CInt(tmpArr(0) + 543))) & "/" & Trim("" & tmpArr(1)) & "/" & Trim("" & tmpArr(2))
        Else
            dtFrom.ValueYMD = backDate
        End If
    End If

    With cboStatus
        .AddItem "== ทุกสถานะ =="
        .AddItem "ยังไม่กำหนด Box"
        .AddItem "กำหนด Box แล้ว"
        .ListIndex = 0
    End With
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

Private Sub Form_Unload(Cancel As Integer)
    VisibleImage
End Sub

Private Sub ClearHeaderIcons(CurrentHeader As Integer)
    Dim i As Integer
    For i = 1 To lvResult.ColumnHeaders.Count
        If lvResult.ColumnHeaders(i).Index <> CurrentHeader Then
            lvResult.ColumnHeaders(i).Icon = Empty
        End If
    Next
End Sub

Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    lvResult.SortKey = ColumnHeader.Index - 1
'    If lvResult.SortOrder = lvwAscending Then
'        lvResult.SortOrder = lvwDescending
'    Else
'        lvResult.SortOrder = lvwAscending
'    End If
    Call ClearHeaderIcons(ColumnHeader.Index)
    Select Case ColumnHeader.Index
        Case 3 'Sorting Date
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortDescending, sortDate)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortDate)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortDate)
            End Select
        Case Else
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortDescending, sortAlpha)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortAlpha)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortAlpha)
            End Select
    End Select
End Sub

Private Sub lvResult_DblClick()
    On Error Resume Next
    Dim tmpS As New BWGSearchManager
    Dim colJob As Collection, colCar As Collection
    Dim JobID As String, JobDataCarID As String, VendorTruckID As String, TimeTableNo As String
    JobID = Replace(Trim(Split(lvResult.SelectedItem.Key, "_")(0)), "'", "")
    JobDataCarID = Replace(Trim(Split(lvResult.SelectedItem.Key, "_")(1)), "'", "") 'lvResult.SelectedItem.SubItems(12)
    VendorTruckID = Replace(Trim(Split(lvResult.SelectedItem.Key, "_")(2)), "'", "") 'lvResult.SelectedItem.SubItems(13)
    TimeTableNo = lvResult.SelectedItem.SubItems(1)
    If Left(TimeTableNo, 1) = "B" Then
        Set colJob = tmpS.JobDataHeaderBoxSearch("JobID= '" & JobID & "'")
        Set colCar = tmpS.JobDataTimeTableBoxSearch("JobDataCarID='" & JobDataCarID & "' AND TimeTableNO = '" & TimeTableNo & "' AND VendorTruckID='" & VendorTruckID & "'")
    Else
        Set colJob = tmpS.JobDataHeaderSearch("JobID= '" & JobID & "'")
        Set colCar = tmpS.JobDataTimeTableSearch("JobDataCarID='" & JobDataCarID & "' AND TimeTableNO = '" & TimeTableNo & "' AND VendorTruckID='" & VendorTruckID & "'")
    End If
    If colJob Is Nothing Then Exit Sub
    If colJob.Count > 0 Then
        Load frmBoxTransferOut
        With frmBoxTransferOut
            If Left(TimeTableNo, 1) = "B" Then
                .TransactionType "JobDataBox"
                .SetDataBoxOpenBox colJob(1), colCar(1)
            Else
                .TransactionType "JobData"
                .SetDataBoxOpen colJob(1), colCar(1)
            End If
            .Show 1
        End With
    End If
    
    Set colJob = Nothing
    Set colCar = Nothing
    Set tmpS = Nothing
End Sub

Private Sub SearchQue()
'    On Error GoTo ErrD
    Dim RS1 As ADODB.Recordset, RS2 As ADODB.Recordset
    Dim sqlStr1 As String, sqlStr2 As String, strCond As String, StatusDesc As String
    Dim tmpArr, CheckAssign As String
    Dim subStr As String
    Dim TotalRecord As Long
    lvResult.ListItems.Clear
    sqlStr1 = "SELECT JobID, JobDataCarID, VendorTruckID, JobNo, TimeTableNo, WorkDate, CompanyCode, CompanyName, CASE WHEN WasteType IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
    sqlStr1 = sqlStr1 & "THEN 'AKP' ELSE CompanyCode END AS DestCompanyCode, CarRegisID, DriverName, TruckTypeDesc, SubTypeDesc," & vbCrLf
    sqlStr1 = sqlStr1 & " AumphurName, ProvinceName, WorkerPerCar , isTrailer, BoxType, BoxCount, AssignedBox, AssignedBoxNo, BoxCanceled,BoxCodeIn" & vbCrLf
    sqlStr1 = sqlStr1 & "From dbo.vw_JobListForBoxAssigned" & vbCrLf
    
    sqlStr2 = "SELECT JobID, JobDataCarID, VendorTruckID, JobNo, TimeTableNo, WorkDate, CompanyCode, CompanyName, CASE WHEN WasteType IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
    sqlStr2 = sqlStr2 & "THEN 'AKP' ELSE CompanyCode END AS DestCompanyCode, CarRegisID, DriverName, TruckTypeDesc, SubTypeDesc," & vbCrLf
    sqlStr2 = sqlStr2 & " AumphurName, ProvinceName, WorkerPerCar , isTrailer, BoxType, BoxCount, AssignedBox, AssignedBoxNo, BoxCanceled,BoxCodeIn" & vbCrLf
    sqlStr2 = sqlStr2 & "From dbo.vw_JobListForBoxAssignedBox" & vbCrLf
    
    If Trim(txtJobNo.Text) <> "" Then
        strCond = "WHERE JobNo Like '%" & Trim(txtJobNo.Text) & "' "
    Else
        strCond = "WHERE (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "') "
        strCond = strCond & " And (CompanyName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or ProvinceName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')" & vbCrLf
        Select Case cboStatus.ListIndex
            Case 1 'Open
                subStr = " (BoxCount>AssignedBox)  AND AssignedBox = 0 And (BoxCodeIn = '' or BoxCodeIn is null )"
            Case 2 'Approve
                subStr = " (BoxCount=AssignedBox)  AND ISNULL(BoxCount,0) > 0"
            Case 3 'Cancel
                subStr = " (BoxCanceled>0)"
        End Select
        
        If cboStatus.ListIndex > 0 Then
            If InStr(1, UCase(strCond), "WHERE") > 0 Then
                strCond = strCond & " AND " & subStr
            Else
                strCond = strCond & "WHERE " & subStr
            End If
        End If
    
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    strCond = strCond & " AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' " & vbCrLf
                    strCond = strCond & "AND JobNo NOT LIKE '%PST%'" & vbCrLf
                ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                    strCond = strCond & " AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' " & vbCrLf
                Else
                    strCond = strCond & " AND CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'"
                    'strCond = strCond & "AND WasteType NOT IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
                    strCond = strCond & "AND JobNo NOT LIKE '%PST%'" & vbCrLf
                End If
            End If
        Else
            If cboCompany.ListIndex >= 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    strCond = strCond & " AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' " & vbCrLf
                    strCond = strCond & "AND JobNo NOT LIKE '%PST%'" & vbCrLf
                ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                    strCond = strCond & " AND CompanyID = '" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "' " & vbCrLf
                Else
                    strCond = strCond & " And CompanyID='" & CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).ID & "'"
                    'strCond = strCond & " AND WasteType NOT IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
                    strCond = strCond & "AND JobNo NOT LIKE '%PST%'" & vbCrLf
                End If
            End If
        End If
        
        If cboCarSubType.ListIndex > 0 Then
            strCond = strCond & " And SubTypeID='" & colCarSubType_(cboCarSubType.ListIndex).ID & "'" & vbCrLf
        End If
    End If
    
    strCond = strCond & "AND ISNULL(BoxType, '') <> '' AND (ISNULL(MenifestCanceled,'N') <> 'Y')"
    curCondition_ = strCond
    sqlStr1 = sqlStr1 & strCond & vbCrLf & "ORDER BY JobNo, TimeTableNo, CarRegisID"
    sqlStr2 = sqlStr2 & strCond & vbCrLf & "ORDER BY JobNo, TimeTableNo, CarRegisID"
    
    Screen.MousePointer = 11
    Set RS1 = New ADODB.Recordset
    Set RS1 = GetRS(sqlStr1, adOpenForwardOnly, adLockReadOnly)
    If RS1.EOF = False Then
        TotalRecord = TotalRecord + RS1.RecordCount
        AddToListView RS1
    End If
    RS1.Close
    Set RS1 = Nothing
    
    Set RS2 = New ADODB.Recordset
    Set RS2 = GetRS(sqlStr2, adOpenForwardOnly, adLockReadOnly)
    If RS2.EOF = False Then
        TotalRecord = TotalRecord + RS2.RecordCount
        AddToListView RS2
    End If
    RS2.Close
    Set RS2 = Nothing
    
    Screen.MousePointer = 0
    frmMain.Stb1.Panels(1).Text = "พบข้อมูล " & Format(TotalRecord, "#,##0") & " รายการ"

    Exit Sub
'ErrD:
'    Screen.MousePointer = 0
'    GetMsg "", vbCritical
'    Exit Sub
End Sub

Private Sub AddToListView(ByVal tmpRec As ADODB.Recordset)
    Dim tmpList As ListItem
    Dim x As Long, Y As Long
    Dim curColor, tripCount As Long
    Do While Not tmpRec.EOF
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpRec!JobID & "_" & tmpRec!JobDataCarID & "_" & tmpRec!VendorTruckID & "_" & Trim(tmpRec!TimeTableNo) & "'", tmpRec!JobNo)
        tmpList.SubItems(1) = Trim("" & tmpRec!TimeTableNo)
        tmpList.SubItems(2) = FormatYMD_to_DMY(Trim("" & tmpRec!WorkDate), "/", "/")
        tmpList.SubItems(3) = Trim("" & tmpRec!CarRegisID)
        tmpList.SubItems(4) = Trim("" & tmpRec!CompanyName)
        tmpList.SubItems(5) = Trim("" & tmpRec!CompanyCode)
        tmpList.SubItems(6) = Trim("" & tmpRec!DestCompanyCode)
        If Trim("" & tmpRec!isTrailer) = "Y" Then
            tmpList.SubItems(7) = Trim("" & tmpRec!TruckTypeDesc) & " (พ่วง), " & Trim("" & tmpRec!SubTypeDesc)
        Else
            tmpList.SubItems(7) = Trim("" & tmpRec!TruckTypeDesc) & ", " & Trim("" & tmpRec!SubTypeDesc)
        End If
        tmpList.SubItems(8) = Trim("" & tmpRec!AumphurName) & ", " & Trim("" & tmpRec!ProvinceName)
        tmpList.SubItems(9) = Trim("" & tmpRec!AssignedBoxNo)
        tmpList.SubItems(10) = Trim("" & tmpRec!BoxType)
        tmpList.SubItems(11) = Trim("" & tmpRec!BoxCount) & " ใบ"
        tmpList.SubItems(12) = Trim("" & tmpRec!AssignedBox) & " ใบ"
        tmpList.SubItems(13) = Trim("" & tmpRec!JobDataCarID)
        tmpList.SubItems(14) = Trim("" & tmpRec!VendorTruckID)
        tmpList.SubItems(15) = Trim("" & tmpRec!BoxCodeIn)
        
        If Trim("" & tmpRec!BoxCount) > 0 And Trim("" & tmpRec!AssignedBox) = 0 Then
            curColor = lbOpen.ForeColor
        ElseIf Trim("" & tmpRec!BoxCount) > Trim("" & tmpRec!AssignedBox) Then
            curColor = lbApprove.ForeColor
        ElseIf Trim("" & tmpRec!BoxCount) = Trim("" & tmpRec!AssignedBox) Then
            curColor = lbApprove.ForeColor
        ElseIf Trim("" & tmpRec!BoxCanceled) > 0 Then
            curColor = lbCancel.ForeColor
        End If
        
        If Trim("" & tmpRec!BoxCanceled) > 0 Then
            curColor = lbCancel.ForeColor
        End If
'        If Trim("" & tmpRec!BoxCount) > 0 And Trim("" & tmpRec!BoxCodeIn) <> "" Then
'            curColor = LbBeIncomplete.ForeColor
'        End If
        
        For Y = 1 To tmpList.ListSubItems.Count
            tmpList.ForeColor = curColor
            tmpList.ListSubItems(Y).ForeColor = curColor
        Next Y

        Set tmpList = Nothing
        'DoEvents
        tmpRec.MoveNext
    Loop

End Sub

Private Sub PopulateCarSubType()
Dim tmpS As New BWGSearchManager
Dim x%
    Set colCarSubType_ = tmpS.TruckSubTypeSearch
    Me.cboCarSubType.Clear
    For x = 1 To colCarSubType_.Count
        cboCarSubType.AddItem colCarSubType_(x).SubTypeDesc
    Next
    cboCarSubType.AddItem "== ทุกประเภทรถ ==", 0
    cboCarSubType.ListIndex = 0
End Sub

