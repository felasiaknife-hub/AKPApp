VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJobForCarAssignList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "คิวงานรอจัดรถ"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12510
   ForeColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9975
   ScaleWidth      =   12510
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
      Height          =   9975
      Left            =   0
      ScaleHeight     =   9945
      ScaleWidth      =   2940
      TabIndex        =   13
      Top             =   0
      Width           =   2970
      Begin MSComctlLib.ListView lvCar 
         Height          =   2445
         Left            =   0
         TabIndex        =   29
         Top             =   5400
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   4313
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ประเภทรถ"
            Object.Width           =   3794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "จำนวน"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.ComboBox CboType 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frJobForCarAssignList.frx":0000
         Left            =   1740
         List            =   "frJobForCarAssignList.frx":0016
         TabIndex        =   26
         Text            =   "--"
         Top             =   3360
         Width           =   1170
      End
      Begin VB.ComboBox cboCustType 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   1545
      End
      Begin VB.CommandButton cmdPrintListJOb 
         Caption         =   "พิมพ์รายการ"
         Height          =   390
         Left            =   1495
         TabIndex        =   22
         Top             =   4095
         Width           =   1395
      End
      Begin DateCtl.Date dtTo 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   2355
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin DateCtl.Date dtFrom 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   1785
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtJobNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Top             =   2880
         Width           =   2760
      End
      Begin VB.ComboBox cboCarSubType 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1215
         Width           =   2700
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   900
         Width           =   2700
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   390
         Left            =   15
         TabIndex        =   11
         Top             =   4095
         Width           =   1395
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   390
         Left            =   15
         TabIndex        =   9
         Top             =   3690
         Width           =   2880
      End
      Begin VB.ComboBox cboCompany 
         Height          =   315
         ItemData        =   "frJobForCarAssignList.frx":005C
         Left            =   180
         List            =   "frJobForCarAssignList.frx":005E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   585
         Width           =   2700
      End
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   270
         Width           =   2670
      End
      Begin MSComctlLib.ListView lvBox 
         Height          =   2130
         Left            =   0
         TabIndex        =   30
         Top             =   7830
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   3757
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ประเภท Box"
            Object.Width           =   3794
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "จำนวน"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.CommandButton cmdViewJobBill 
         Caption         =   "ดูข้อมูลใบ Manifest"
         Height          =   435
         Left            =   60
         TabIndex        =   10
         Top             =   7800
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   240
         Index           =   5
         Left            =   120
         Top             =   5040
         Width           =   195
      End
      Begin VB.Label LblVIPCR 
         AutoSize        =   -1  'True
         Caption         =   "VIP CR"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   405
         TabIndex        =   28
         Top             =   5085
         Width           =   555
      End
      Begin VB.Label LblBME 
         AutoSize        =   -1  'True
         Caption         =   "Better ME"
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   405
         TabIndex        =   27
         Top             =   4545
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภท Order"
         Height          =   195
         Index           =   5
         Left            =   1740
         TabIndex        =   25
         Top             =   3165
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ประเภทลูกค้า :"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   23
         Top             =   3165
         Width           =   1020
      End
      Begin VB.Label LbcarAP 
         BackColor       =   &H80000005&
         Caption         =   "รถ AP"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   405
         TabIndex        =   21
         Top             =   7335
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   3
         Left            =   120
         Top             =   7335
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบสั่งงาน :"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   2685
         Width           =   1080
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   390
         TabIndex        =   7
         Top             =   7905
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closed"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   7665
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblNormal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รถโรงงาน"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   405
         TabIndex        =   19
         Top             =   4800
         Width           =   705
      End
      Begin VB.Label lblDTeam 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DTeam"
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   1350
         TabIndex        =   18
         Top             =   7560
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   120
         Top             =   7500
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   120
         Top             =   7665
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   120
         Top             =   4770
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   120
         Top             =   4515
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
         Top             =   2115
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
         Top             =   1545
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
         Top             =   30
         Width           =   1530
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   3000
      TabIndex        =   12
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   4577
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImgSorted"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   22
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "เลขที่ใบสั่งงาน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "ลำดับ"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "วันที่ลง Order"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "วันที่อัพเดท Order"
         Object.Width           =   2734
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "วันที่ขน"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ลูกค้า"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "บริษัท"
         Object.Width           =   1162
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ปลายทาง"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "รถที่ใช้ขน"
         Object.Width           =   3678
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "สถานที่ขน"
         Object.Width           =   4075
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "จำนวนเที่ยว"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "คนงาน"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "กำหนดรถแล้ว"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "จำนวน Manifest"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "CarRem0"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "CarRam1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "CarRam2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "CarRam3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Vendor"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "OrderType"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "CarRem"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "ตรวจสอบ"
         Object.Width           =   2540
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
            Picture         =   "frJobForCarAssignList.frx":0060
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frJobForCarAssignList.frx":05FA
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
Attribute VB_Name = "frmJobForCarAssignList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCondition_ As String
Dim colCarSubType_ As Collection

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrintListJOb_Click()
Dim tmpRec As ADODB.Recordset
Dim tmpList As ListItem
Dim x As Long, y As Long
Dim sqlStr As String, strCond As String, StatusDesc As String
Dim tmpArr
Dim curColor, tripCount As Long
Dim subStr As String
    sqlStr = "SELECT     JobID, JobDataCarID, JobNo, WorkDate, CompanyName, CompanyCode, CASE WHEN WasteType IN ('HIC', 'HBI', 'NIC', 'NBI', 'BBH', 'BBL')" & vbCrLf
    sqlStr = sqlStr & "                      THEN 'AKP' ELSE CompanyCode END AS DestCompanyCode, TruckTypeDesc, SubTypeDesc, AumphurName, ProvinceName, TruckCount, AssignedCount," & vbCrLf
    sqlStr = sqlStr & "                      WorkerPerCar , MenifestCount, isTrailer, ISNULL(DTeam, 'N') AS DTeam, IssuedDate, IssuedTime," & vbCrLf
      sqlStr = sqlStr & "                    SUBSTRING(CarRem,1,4) as CarRem1,SUBSTRING(CarRem,1,7) as CarRem2,SUBSTRING(CarRem,1,6) as CarRem3" & vbCrLf
    sqlStr = sqlStr & "From dbo.vw_JobListForCarAssigned" & vbCrLf
    If Trim(txtJobNo.Text) <> "" Then
        strCond = " WHERE JobNo Like '%" & Trim(txtJobNo.Text) & "' "
    Else
        strCond = " WHERE (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "') "
        strCond = strCond & " AND (CompanyName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' OR ProvinceName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')" & vbCrLf
        Select Case cboStatus.ListIndex
            Case 1 'Open
                subStr = " (TruckCount>AssignedCount) "
            Case 2 'Approve
                subStr = " (TruckCount=AssignedCount) "
        End Select
        If cboStatus.ListIndex > 0 Then
            If InStr(1, UCase(strCond), "WHERE") > 0 Then
                strCond = strCond & " AND " & subStr
            Else
                strCond = strCond & " WHERE " & subStr
            End If
        End If
    Dim compID As String
    Dim CompCode As String
    If cboCompany.ListIndex = 1 Then
    CompCode = "AKP"
    compID = "776EDA5830F744A6B3551348D851BC22"
    Else
     CompCode = ""
    compID = ""
    End If
    
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    strCond = strCond & " AND (CompanyID = '" & compID & "') " & vbCrLf
                    strCond = strCond & " AND JobNo NOT LIKE '%PST%'" & vbCrLf
                ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                    strCond = strCond & " AND (CompanyID = '" & compID & "') " & vbCrLf
                Else
                    strCond = strCond & " AND CompanyID='" & compID & "'"
                    'strCond = strCond & " AND WasteType NOT IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
                    strCond = strCond & " AND JobNo NOT LIKE '%PST%'" & vbCrLf
                End If
            End If
        Else
            If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                strCond = strCond & " AND (CompanyID = '" & compID & "') " & vbCrLf
                strCond = strCond & " AND JobNo NOT LIKE '%PST%'" & vbCrLf
            ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                strCond = strCond & " AND (CompanyID = '" & compID & "') " & vbCrLf
            Else
                strCond = strCond & " AND CompanyID='" & compID & "'"
                'strCond = strCond & " AND WasteType NOT IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
                strCond = strCond & " AND JobNo NOT LIKE '%PST%'" & vbCrLf
            End If
        End If
        
        If cboCarSubType.ListIndex > 0 Then
            strCond = strCond & " AND SubTypeID='" & colCarSubType_(cboCarSubType.ListIndex).id & "'" & vbCrLf
        End If
    End If
    sqlStr = sqlStr & strCond & " ORDER BY WorkDate Asc, JobNo Desc"
    PrintMain 0, sqlStr
    
End Sub
Private Sub PrintMain(printType As Integer, tmpSql As String)
    With rpt1
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\JobListAssinged.rpt"
'        Dim strstring As String
'        strstring = "ในช่วงวันที่ขนตั้งแต่ " & dtFrom.ValueDMY & " ถึงวันที่ " & dtTo.ValueDMY & Chr(34) & cboCompany.Text & Chr(34)
'        .Formulas(0) = "RptHeader=" & strstring
        .WindowTitle = "รายงานคิวงานรอจัดรถ"
        .WindowShowCloseBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .SQLQuery = tmpSql
        .Destination = printType
        .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
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
        .AddItem "ยังไม่กำหนดรถ"
        .AddItem "กำหนดรถแล้ว"
        .ListIndex = 1
    End With
    With cboCustType
        .AddItem "== ประเภทลูกค้า =="
        .AddItem "(C) ,(s)"
        .AddItem "(M) ,(MS)"
        .AddItem "(B) ,(E)"
        .AddItem "VIP CR"
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

Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    lvResult.SortKey = ColumnHeader.Index - 1
'    If lvResult.SortOrder = lvwAscending Then
'        lvResult.SortOrder = lvwDescending
'    Else
'        lvResult.SortOrder = lvwAscending
'    End If
    Call ClearHeaderIcons(ColumnHeader.Index)
    Select Case ColumnHeader.Index
        Case 3
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

Private Sub ClearHeaderIcons(CurrentHeader As Integer)
    Dim i As Integer
    For i = 1 To lvResult.ColumnHeaders.Count
        If lvResult.ColumnHeaders(i).Index <> CurrentHeader Then
            lvResult.ColumnHeaders(i).Icon = Empty
        End If
    Next
End Sub

Private Sub lvResult_DblClick()
On Error Resume Next
Dim JobID As String
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    JobID = Replace(Trim(Split(lvResult.SelectedItem.Key, "_")(0)), "'", "")
    Set tmpCol = tmpS.JobDataHeaderSearch("JobID='" & JobID & "'")
    If tmpCol Is Nothing Then Exit Sub
    If tmpCol.Count > 0 Then
        Load frmJobCarAssign
        With frmJobCarAssign
            .SetDataJobOpen tmpCol(1)
            .Show 1
        End With
    End If
    Set tmpCol = Nothing
    Set tmpS = Nothing
End Sub

Private Sub SearchQue()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim tmpList As ListItem
Dim ListCar As ListItem
Dim ListAKPBox As ListItem
Dim x As Long, y As Long
Dim sqlStr As String, strCond As String, StatusDesc As String
Dim tmpArr
Dim curColor, tripCount As Long
Dim subStr As String
    sqlStr = "SELECT     JobID, JobDataCarID, JobNo, WorkDate,NumberRow, CompanyName + ' [' +  CustomerCode +  ']'  + '-' + StaffFName + ' ' + StaffLName as CompanyName, CompanyCode, CASE WHEN WasteType IN ('HIC', 'HBI', 'NIC', 'NBI', 'BBH', 'BBL')" & vbCrLf
    sqlStr = sqlStr & "                      THEN 'AKP' ELSE CompanyCode END AS DestCompanyCode, TruckTypeDesc, SubTypeDesc, AumphurName, ProvinceName, TruckCount, AssignedCount," & vbCrLf
    sqlStr = sqlStr & "                      WorkerPerCar , MenifestCount, isTrailer, ISNULL(DTeam, 'N') AS DTeam, IssuedDate, IssuedTime, carrem, " & vbCrLf
      sqlStr = sqlStr & "                    SUBSTRING(CarRem,1,4) as CarRem1,SUBSTRING(CarRem,1,7) as CarRem2,SUBSTRING(CarRem,1,6) as CarRem3,VendorName,OrderType,isVIPCR,NewCustJob,LastJobUpdate,IsChk" & vbCrLf
    sqlStr = sqlStr & "From dbo.vw_JobListForCarAssigned" & vbCrLf
    If Trim(txtJobNo.Text) <> "" Then
        strCond = " WHERE JobNo Like '%" & Trim(txtJobNo.Text) & "' "
    Else
        strCond = " WHERE (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "') "
        strCond = strCond & " AND (CompanyName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' OR ProvinceName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')" & vbCrLf
        Select Case cboStatus.ListIndex
            Case 1 'Open
                subStr = " (TruckCount>AssignedCount)  "
            Case 2 'Approve
                subStr = " (TruckCount=AssignedCount) "
        End Select
        If cboStatus.ListIndex > 0 Then
            If InStr(1, UCase(strCond), "WHERE") > 0 Then
                strCond = strCond & " AND " & subStr
            Else
                strCond = strCond & " WHERE " & subStr
            End If
        End If
        
            Dim comp As String
                If cboCompany.ListIndex = 0 Then
                    comp = "776EDA5830F744A6B3551348D851BC22"
                Else
                    comp = "776EDA5830F744A6B3551348D851BC22"
                End If
    
        If cboCompany.ListCount > 1 Then
            If cboCompany.ListIndex > 0 Then
                If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                    strCond = strCond & " AND (CompanyID = '" & comp & "') " & vbCrLf
                    strCond = strCond & " AND JobNo NOT LIKE '%PST%'" & vbCrLf
                ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                    strCond = strCond & " AND (CompanyID = '" & comp & "') " & vbCrLf
                Else
                    strCond = strCond & " AND CompanyID='" & comp & "'"
                    'strCond = strCond & " AND WasteType NOT IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
                    strCond = strCond & " AND JobNo NOT LIKE '%PST%'" & vbCrLf
                End If
            End If
        Else
            If CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "AKP" Then
                strCond = strCond & " AND (CompanyID = '" & comp & "') " & vbCrLf
                strCond = strCond & " AND JobNo NOT LIKE '%PST%'" & vbCrLf
            ElseIf CurrentUser.WorkCompanies(cboCompany.ListIndex + cboIndex).CompanyCode = "PST" Then
                strCond = strCond & " AND (CompanyID = '" & comp & "') " & vbCrLf
            Else
                strCond = strCond & " AND CompanyID='" & comp & "'"
                'strCond = strCond & " AND WasteType NOT IN ('HIC', 'HBI', 'BBH', 'BBL', 'NIC', 'NBI')" & vbCrLf
                strCond = strCond & " AND JobNo NOT LIKE '%PST%'" & vbCrLf
            End If
        End If
        
        If cboCarSubType.ListIndex > 0 Then
            strCond = strCond & " AND SubTypeID='" & colCarSubType_(cboCarSubType.ListIndex).id & "'" & vbCrLf
        End If
    End If
    If cboCustType.ListIndex = 1 Then
            strCond = strCond & " AND CustTypeFlag in ('C','S') "
        ElseIf cboCustType.ListIndex = 2 Then
            strCond = strCond & " AND CustTypeFlag in ('M','MS') "
        ElseIf cboCustType.ListIndex = 3 Then
            strCond = strCond & " AND CustTypeFlag in ('B','E') "
        ElseIf cboCustType.ListIndex = 4 Then
            strCond = strCond & " AND isVIPCR = 'Y' "
    End If
          
   If cboType.ListIndex = 1 Then
        strCond = strCond & " AND OrderType = 'B' "
    ElseIf cboType.ListIndex = 2 Then
        strCond = strCond & " AND OrderType = 'K' "
    ElseIf cboType.ListIndex = 3 Then
        strCond = strCond & " AND OrderType = 'P' "
    ElseIf cboType.ListIndex = 4 Then
        strCond = strCond & " AND OrderType = 'U' "
    ElseIf cboType.ListIndex = 5 Then
        strCond = strCond & " AND OrderType = 'C' "
    ElseIf cboType.ListIndex = 6 Then
        strCond = strCond & " AND OrderType = 'S' "
    ElseIf cboType.ListIndex = 7 Then
        strCond = strCond & " AND OrderType = 'W' "
    End If
        
    curCondition_ = strCond
    Screen.MousePointer = 11
    Set tmpRec = New ADODB.Recordset
    Set tmpRec = GetRS(sqlStr & strCond & " ORDER BY IssuedDate,WorkDate Desc", adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ"
        Exit Sub
    End If
    lvResult.ListItems.Clear
    If tmpRec Is Nothing Then
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ  จำนวน 0 เที่ยว"
'        MsgBox "ไม่พบข้อมูล", vbInformation, "ผลการค้นหา"
        Exit Sub
    End If
    If tmpRec.EOF = False Then
        Do While Not tmpRec.EOF
            Set tmpList = lvResult.ListItems.Add(, "'" & tmpRec!JobID & "_" & tmpRec!JobDataCarID & "'", "" & tmpRec!JobNo)
             tmpList.SubItems(1) = tmpRec!NumberRow
            tmpList.SubItems(2) = Format(Trim("" & tmpRec!IssuedDate), "dd/MM/yyyy") & " " & Format(Trim("00" & tmpRec!IssuedTime), "HH:mm:ss")                    '// ORDER DATE
            tmpList.SubItems(3) = Format(Trim("" & tmpRec!LastJobUpdate), "dd/MM/yyyy") & " " & Format(Trim("00" & tmpRec!LastJobUpdate), "HH:mm:ss")                    '// Last Update
            tmpList.SubItems(4) = FormatYMD_to_DMY(Trim("" & tmpRec!WorkDate), "/", "/")
            tmpList.SubItems(5) = tmpRec!CompanyName
            tmpList.SubItems(6) = Trim("" & tmpRec!CompanyCode)
            tmpList.SubItems(7) = Trim("" & tmpRec!DestCompanyCode)
            If Trim("" & tmpRec!isTrailer) = "Y" Then
                tmpList.SubItems(8) = Trim("" & tmpRec!TruckTypeDesc) & " (พ่วง), " & Trim("" & tmpRec!SubTypeDesc)
            Else
                tmpList.SubItems(8) = Trim("" & tmpRec!TruckTypeDesc) & ", " & Trim("" & tmpRec!SubTypeDesc)
            End If
            tmpList.SubItems(9) = Trim("" & tmpRec!AumphurName) & ", " & Trim("" & tmpRec!ProvinceName)
            tmpList.SubItems(10) = Trim("" & tmpRec!TruckCount) & " เที่ยว"
            If IsNumeric(Trim("" & tmpRec!TruckCount)) Then
                tripCount = tripCount + CDbl(Trim("" & tmpRec!TruckCount))
            End If
            tmpList.SubItems(11) = Trim("" & tmpRec!WorkerPerCar) & " คน"
            tmpList.SubItems(12) = Trim("" & tmpRec!AssignedCount) & " เที่ยว"
            tmpList.SubItems(13) = Trim("" & tmpRec!MenifestCount) & " ใบ"
            tmpList.SubItems(14) = Trim("" & tmpRec!CarRem)
            tmpList.SubItems(15) = Trim("" & tmpRec!CarRem1)
            tmpList.SubItems(16) = Trim("" & tmpRec!CarRem2)
            tmpList.SubItems(17) = Trim("" & tmpRec!CarRem3)
            tmpList.SubItems(18) = Trim("" & tmpRec!VendorName)
            If Trim("" & tmpRec!OrderType) = "B" Then
                tmpList.SubItems(19) = "งาน BOI"
            ElseIf Trim("" & tmpRec!OrderType) = "K" Then
                 tmpList.SubItems(19) = "งานกรมศุล"
            ElseIf Trim("" & tmpRec!OrderType) = "P" Then
                 tmpList.SubItems(19) = "งานสินค้าทำลาย"
            ElseIf Trim("" & tmpRec!OrderType) = "U" Then
                 tmpList.SubItems(19) = "'งานสินค้าส่งออก"
'            ElseIf Trim("" & tmpRec!OrderType) = "C" Then
'                 tmpList.SubItems(18) = "งานลูกค้าใหม่"
            ElseIf Trim("" & tmpRec!OrderType) = "S" Then
                 tmpList.SubItems(19) = "งานสรรพากร"
            ElseIf Trim("" & tmpRec!OrderType) = "W" Then
                 tmpList.SubItems(19) = "Prowaste"
            Else
                tmpList.SubItems(19) = "-"
            End If
            
            If Trim("" & tmpRec!NewCustJob) = "Y" Then
                 tmpList.SubItems(19) = "งานลูกค้าใหม่"
            End If
            
            tmpList.SubItems(20) = Trim("" & tmpRec!CarRem)
            
            If Trim("" & tmpRec!IsChk) = "Y" Then
                 tmpList.SubItems(21) = "Y"
            Else
                 tmpList.SubItems(21) = "N"
            End If
            If Trim("" & tmpRec!DTeam) = "Y" Then
                curColor = lblDTeam.ForeColor
                tmpList.ForeColor = lblDTeam.ForeColor
            Else
                curColor = LbcarAP.ForeColor
                tmpList.ForeColor = LbcarAP.ForeColor
            End If
            If Trim("" & tmpRec!TruckTypeDesc) = "รถโรงงาน" Then
                curColor = lblNormal.ForeColor
                tmpList.ForeColor = lblNormal.ForeColor
            End If
            
'            If Trim("" & tmpRec!CarRem3) <> "รถโรงง" Or ("" & tmpRec!CarRem2) <> "รถโรงงา" Then
'                curColor = LbcarAP.ForeColor
'                tmpList.ForeColor = LbcarAP.ForeColor
'            End If
            If Trim("" & tmpRec!CarRem3) = "งานBME" Then
                curColor = LblBME.ForeColor
                tmpList.ForeColor = LblBME.ForeColor
            End If
            
            If Trim("" & tmpRec!isVIPCR) = "Y" Then
                curColor = LblVIPCR.ForeColor
                tmpList.ForeColor = LblVIPCR.ForeColor
            End If
            
            For y = 1 To tmpList.ListSubItems.Count
                tmpList.ListSubItems(y).ForeColor = curColor
            Next
            
            Set tmpList = Nothing
            'DoEvents
            tmpRec.MoveNext
        Loop 'Next
    End If
    
    ' แสดงจำนวนรถคงเหลือจากการลง Order
    Dim tSql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
    Dim CU As Integer
    tSql = "SELECT CarDailyID,SubtypeID,SubTypeDesc,CarUse,CarUsed FROM tbCarDaily where workdate='" & dtTo.ValueYMD & "' " & vbCrLf
    tSql = tSql & "order by SubTypeDesc"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
    lvCar.ListItems.Clear
        Do While Not rs.EOF
        CU = "0"
        If rs!CarUsed <> "Null" Then
            CU = rs!CarUsed
            Else
            CU = 0
        End If
            Set ListCar = lvCar.ListItems.Add(, rs!CarDailyID)
            ListCar.SubItems(1) = rs!SubTypeDesc
            ListCar.SubItems(2) = rs!CarUse
            For i = 1 To ListCar.ListSubItems.Count
                If ListCar.SubItems(2) <= 0 Then
                    ListCar.ListSubItems(i).ForeColor = vbRed
                End If
            Next
            Set ListCar = Nothing
        rs.MoveNext
        Loop

'แสดงจำนวน BOX คงเหลือจากการลง Order
    Dim SqlBox As String
    Dim RSBox As New ADODB.Recordset
    Dim j As Integer
    SqlBox = "SELECT * FROM tbAKPBoxUse where workdate='" & dtTo.ValueYMD & "' " & vbCrLf
    SqlBox = SqlBox & "order by boxcode"
    Set RSBox = GetRS(SqlBox, adOpenForwardOnly, adLockOptimistic)
    lvBox.ListItems.Clear
        Do While Not RSBox.EOF
            Set ListAKPBox = lvBox.ListItems.Add(, RSBox!id)
            ListAKPBox.SubItems(1) = RSBox!boxname
            ListAKPBox.SubItems(2) = RSBox!boxUse
            For i = 1 To ListAKPBox.ListSubItems.Count
                If ListAKPBox.SubItems(2) <= 0 Then
                    ListAKPBox.ListSubItems(i).ForeColor = vbRed
                End If
            Next
            Set ListAKPBox = Nothing
        RSBox.MoveNext
        Loop

    
    Screen.MousePointer = 0
    frmMain.Stb1.Panels(1).Text = "พบข้อมูล " & Format(tmpRec.RecordCount, "#,##0") & " รายการ  จำนวน " & tripCount & " เที่ยว"
    tmpRec.Close
    Set tmpRec = Nothing
    Exit Sub
ErrD:
    Screen.MousePointer = 0
    GetMsg "", vbCritical
    Exit Sub
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

