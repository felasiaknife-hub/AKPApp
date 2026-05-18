VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenifestStatusList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "สถานะใบ Menifest - F062"
   ClientHeight    =   9975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13410
   ForeColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9975
   ScaleWidth      =   13410
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Rpt3 
      Left            =   12420
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      ItemData        =   "frMenifestStatusList.frx":0000
      Left            =   3390
      List            =   "frMenifestStatusList.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   30
      Visible         =   0   'False
      Width           =   5205
   End
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
      ScaleWidth      =   2625
      TabIndex        =   14
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton CmdPrintCont 
         Caption         =   "พิมพ์ใบส่งของ"
         Height          =   435
         Left            =   90
         TabIndex        =   32
         Top             =   6885
         Width           =   2475
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   435
         Left            =   90
         TabIndex        =   31
         Top             =   7335
         Visible         =   0   'False
         Width           =   2475
      End
      Begin VB.CommandButton CmdTrans 
         Caption         =   "ใบกำกับการขนส่ง"
         Height          =   435
         Left            =   90
         TabIndex        =   30
         Top             =   5985
         Width           =   2475
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   3045
         Width           =   2295
      End
      Begin VB.CommandButton cmdChangeTripNo 
         Caption         =   "เปลี่ยนเลขที่ใบคุม"
         Height          =   435
         Left            =   90
         TabIndex        =   11
         Top             =   6435
         Width           =   2475
      End
      Begin VB.TextBox txtJobNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   3600
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-"
         Height          =   255
         Left            =   780
         TabIndex        =   25
         Top             =   3930
         Width           =   555
      End
      Begin VB.CommandButton Command3 
         Caption         =   "+"
         Height          =   255
         Left            =   180
         TabIndex        =   24
         Top             =   3930
         Width           =   555
      End
      Begin VB.CommandButton Command2 
         Caption         =   "พิมพ์ใบคุม"
         Height          =   435
         Left            =   90
         TabIndex        =   10
         Top             =   5550
         Width           =   2475
      End
      Begin VB.CommandButton Command1 
         Caption         =   "พิมพ์ใบ Manifest /ใบส่งมอบงาน"
         Height          =   435
         Left            =   90
         TabIndex        =   9
         Top             =   5085
         Width           =   2475
      End
      Begin VB.TextBox txtKey2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Text            =   "*"
         Top             =   960
         Width           =   2295
      End
      Begin DateCtl.Date dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   3
         Top             =   1905
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.CommandButton cmdViewJobBill 
         Caption         =   "ดูข้อมูลใบ Manifest"
         Height          =   435
         Left            =   90
         TabIndex        =   8
         Top             =   4635
         Width           =   2475
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   90
         TabIndex        =   12
         Top             =   7785
         Width           =   2475
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   90
         TabIndex        =   7
         Top             =   4200
         Width           =   2475
      End
      Begin VB.ComboBox cboJobStatus 
         Height          =   315
         ItemData        =   "frMenifestStatusList.frx":0004
         Left            =   180
         List            =   "frMenifestStatusList.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   2325
      End
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   360
         Width           =   2295
      End
      Begin DateCtl.Date dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   4
         Top             =   2475
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Line Line2 
         X1              =   180
         X2              =   2490
         Y1              =   2820
         Y2              =   2820
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   29
         Top             =   2850
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบสั่งงาน,เลขที่ใบคุม :"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   26
         Top             =   3405
         Width           =   1890
      End
      Begin VB.Line Line1 
         X1              =   180
         X2              =   2490
         Y1              =   3915
         Y2              =   3915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor, คนขับ, ทะเบียนรถ:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   720
         Width           =   1965
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   450
         TabIndex        =   22
         Top             =   9435
         Width           =   480
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปิดงานแล้ว"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   450
         TabIndex        =   21
         Top             =   9135
         Width           =   780
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พิมพ์ Manifest แล้ว"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   450
         TabIndex        =   20
         Top             =   8835
         Width           =   1335
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "งานใหม่"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   450
         TabIndex        =   19
         Top             =   8535
         Width           =   570
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   180
         Top             =   9405
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   180
         Top             =   9105
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   180
         Top             =   8805
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   180
         Top             =   8505
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   18
         Top             =   2235
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ปฏิบัติงานตั้งแต่"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   17
         Top             =   1665
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลข Manifest, ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   15
         Top             =   120
         Width           =   1650
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2610
      TabIndex        =   13
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImgSorted"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   21
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Manifest"
         Object.Width           =   4057
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ลูกค้า"
         Object.Width           =   2036
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "เลขที่ใบสั่งงาน"
         Object.Width           =   2884
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "รถที่ใช้ขน"
         Object.Width           =   2618
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Vendor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "คนขับ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "วันที่ขน"
         Object.Width           =   1984
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "วันที่อนุมัติ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "สถานะ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "เลขที่ใบคุม"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ผู้ปิดงาน"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ผู้อนุมัติการปิดงาน"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "เบอร์คนขัย"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Sale"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "GPS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "วิธีส่ง"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "เรท(พื้นฐาน)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "เรท(จริง)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "งานลูกค้าใหม่"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Object.Width           =   0
      EndProperty
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrinterName     =   "ManifestPrinter"
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin Crystal.CrystalReport rpt2 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrinterName     =   "ManifestPrinter"
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComctlLib.ImageList ImgSorted 
      Left            =   7920
      Top             =   3960
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
            Picture         =   "frMenifestStatusList.frx":0008
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frMenifestStatusList.frx":05A2
            Key             =   "down"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เครื่องพิมพ์"
      Height          =   195
      Index           =   5
      Left            =   2490
      TabIndex        =   27
      Top             =   90
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อรุ่น,ลาย หรือยี่ห้อ :"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   16
      Top             =   4950
      Width           =   1440
   End
End
Attribute VB_Name = "frmMenifestStatusList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCondition_ As String
Dim colControlDoc_ As Collection
Dim PrintSelected As Boolean
Dim Lbl As String
Dim SaleStaff_ As BWGCompanyStaff
Private Sub cmdChangeTripNo_Click()
Dim tmpStr As String
Dim selTripNo As String
    tmpStr = CurrentUser.VerifyAccessRole("frmChangeTripNo", AccessRole)
    If Trim(tmpStr) <> "" Then
        MsgBox tmpStr, vbExclamation
        Exit Sub
    End If
    If lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกใบ Manifest ที่ต้องการแก้ไขเลขที่ใบคุม", vbExclamation
        Exit Sub
    End If
    Load frmChangeTripNo
    With frmChangeTripNo
        selTripNo = Trim(lvResult.SelectedItem.SubItems(9))
        .SetTripNo selTripNo
        .JobDataTimeTable "tbJobDataTimeTable"
        .SetCallerForm Me
        .Show 1
        If isExist(colControlDoc_, "'" & selTripNo & "'") Then
            colControlDoc_("'" & selTripNo & "'").TimeTableNo = lvResult.SelectedItem.SubItems(9)
        End If
    End With
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub CmdExcel_Click()
'    Dim objExcel As New Excel.Application
'
'    Dim objExcelSheet As Excel.Worksheet
'    Dim col, row As Integer
'
'    If lvResult.ListItems.Count > 0 Then
'        objExcel.Workbooks.Add
'        Set objExcelSheet = objExcel.Worksheets.Add
'        For col = 1 To lvResult.ColumnHeaders.Count
'            objExcelSheet.Cells(1, col).Value = lvResult.ColumnHeaders(col)
'        Next
'        For row = 2 To lvResult.ListItems.Count + 1
'           For col = 1 To lvResult.ColumnHeaders.Count
'           If col = 1 Then
'                   objExcelSheet.Cells(row, col).Value = lvResult.ListItems(row - 1).Text
'           Else
'                   objExcelSheet.Cells(row, col).Value = lvResult.ListItems(row - 1).SubItems(col - 1)
'           End If
'           Next
'        Next row
'        objExcelSheet.Columns.AutoFit
'
'        objExcel.Visible = True
'
'    Else
'        MsgBox "No data to export", vbInformation, Me.Caption
'    End If
End Sub

Private Sub CmdPrintCont_Click()
Dim x%, iCount%
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีรายการใบ Manifest ที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    If colControlDoc_ Is Nothing Then
        MsgBox "ไม่มีรายการใบกำกับการขนส่งที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    If colControlDoc_.Count = 0 Then
        MsgBox "ไม่มีรายการใบกำกับการขนส่งที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    Dim tmpRec As ADODB.Recordset
    PrintSelected = False
    Dim sqlStr As String
    Dim tmpCol As Collection
    Dim checkCount As Integer
    checkCount = 0
    For x = 1 To lvResult.ListItems.Count
        If lvResult.ListItems(x).Checked Then
            sqlStr = "SELECT * FROM vw_JobDataCarUsedSearch WHERE jobdatacarid= '" & lvResult.ListItems(x).SubItems(20) & "'"
            Set tmpRec = New ADODB.Recordset
            Set tmpCol = New Collection
            Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
                With RPT3
                    .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
                    .ReportFileName = App.Path & "\reports\Deliver.rpt"
                    .SQLQuery = sqlStr
                    .Destination = crptToPrinter
                    .PrintReport
                    checkCount = checkCount + 1
                If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
               End With
       End If
    Next
        PrintSelected = False
        MsgBox "พิมพ์ข้อมูลใบส่งของเรียบร้อยแล้ว จำนวน " & checkCount & " รายการ", vbInformation
End Sub
Private Sub cmdSearch_Click()
    SearchQue
    If lvResult.ListItems.Count = 0 Then MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
End Sub
Private Sub CmdTrans_Click()
Dim x%, iCount%
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีรายการใบ Manifest ที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    If colControlDoc_ Is Nothing Then
        MsgBox "ไม่มีรายการใบกำกับการขนส่งที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    If colControlDoc_.Count = 0 Then
        MsgBox "ไม่มีรายการใบกำกับการขนส่งที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    For x = 1 To colControlDoc_.Count
    Dim tmpRec As ADODB.Recordset
    Dim tmpRegisH  As ADODB.Recordset
    Dim tmpRegisT  As ADODB.Recordset
    PrintSelected = False
    Dim sqlStr As String
    Dim tmpCol As Collection
        sqlStr = "SELECT * FROM vw_MenifestDetInTimeNo WHERE isCanceled<> 'Y' and  TimeTableNo='" & colControlDoc_(x).TimeTableNo & "'   ORDER BY companyName    "
        Set tmpRec = New ADODB.Recordset
        Set tmpCol = New Collection
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        With rpt2
            .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
            .ReportFileName = App.Path & "\reports\TripMenifestDet.rpt"
              If tmpRec(22) = "รถ Roll off Truck- พ่วง" Or tmpRec(0) = "รถ Lugger Box - พ่วง" Then
                Dim sqlrs1 As String
                sqlrs1 = " SELECT distinct CarRegisID,TruckTypeDesc,HeadorTrail FROM vw_MenifestDetInTimeNo WHERE isCanceled<> 'Y' and HeadorTrail <> 'T'  and  TimeTableNo='" & colControlDoc_(x).TimeTableNo & "'   ORDER BY HeadorTrail  "
                Set tmpRegisH = New ADODB.Recordset
               Set tmpRegisH = GetRS(sqlrs1, adOpenForwardOnly, adLockReadOnly)
                    If tmpRegisH.EOF Then
                        .Formulas(0) = "CarRegisH=" & Chr(34) & "" & Chr(34)
                        .Formulas(1) = "TypeHdesc=" & Chr(34) & "" & Chr(34)
                        .Formulas(2) = "CarRegisT=" & Chr(34) & "" & Chr(34)
                        .Formulas(3) = "TypeTdesc=" & Chr(34) & "" & Chr(34)
                        .Formulas(8) = "APName=" & Chr(34) & CurrentUser.StaffFName & " " & CurrentUser.StaffLName & Chr(34)
                    Else
                        .Formulas(0) = "CarRegisH=" & Chr(34) & tmpRegisH(0) & Chr(34)
                        .Formulas(1) = "TypeHdesc=" & Chr(34) & tmpRegisH(1) & Chr(34)
                        .Formulas(8) = "APName=" & Chr(34) & CurrentUser.StaffFName & " " & CurrentUser.StaffLName & Chr(34)
                          Dim sqlrs2 As String
                        sqlrs2 = " SELECT distinct CarRegisID,TruckTypeDesc,HeadorTrail FROM vw_MenifestDetInTimeNo WHERE isCanceled<> 'Y' and HeadorTrail <> 'H'  and  TimeTableNo='" & colControlDoc_(x).TimeTableNo & "'   ORDER BY HeadorTrail  "
                        Set tmpRegisT = New ADODB.Recordset
                        Set tmpRegisT = GetRS(sqlrs2, adOpenForwardOnly, adLockReadOnly)
                        If tmpRegisT.EOF Then
                            .Formulas(2) = "CarRegisT=" & Chr(34) & "" & Chr(34)
                            .Formulas(3) = "TypeTdesc=" & Chr(34) & "" & Chr(34)
                        End If
                    End If
              Else
                .Formulas(0) = "CarRegisH=" & Chr(34) & tmpRec(25) & Chr(34)
                .Formulas(1) = "TypeHdesc=" & Chr(34) & tmpRec(22) & Chr(34)
                .Formulas(2) = "CarRegisT=" & Chr(34) & "-" & Chr(34)
                .Formulas(3) = "TypeTdesc=" & Chr(34) & "-" & Chr(34)
                .Formulas(8) = "APName=" & Chr(34) & CurrentUser.StaffFName & " " & CurrentUser.StaffLName & Chr(34)
              End If
        .SQLQuery = sqlStr
            .Destination = crptToPrinter
            .PrintReport
        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
       End With
    Next
        PrintSelected = False
     MsgBox "พิมพ์ใบกำกับการขนส่งเรียบร้อยแล้ว", vbExclamation
End Sub

Private Sub cmdViewJobBill_Click()
    lvResult_DblClick
End Sub
Private Sub Command1_Click()
Dim x%, iCount%, tmpItem As BWGMenifest, tmpCon As ADODB.Connection
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีรายการใบ Manifest ที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    PrintSelected = False
    
    For x = 1 To lvResult.ListItems.Count
        If lvResult.ListItems(x).Checked Then
            Set tmpCol = tmpS.MenifestSearch("MenifestID=" & lvResult.ListItems(x).Key)
            If tmpCol.Count > 0 Then
                iCount = iCount + 1
                PrintMain tmpCol(1)
            End If
            Set tmpCol = Nothing
        End If
    Next
    
    PrintSelected = False
    Set tmpCon = Nothing
    Set tmpS = Nothing
    If iCount = 0 Then
        MsgBox "กรุณาเลือกรายการใบ Manifest ที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    'เวลาพิมพ์อย่าลืมเช็คว่าจะใช้ที่อยู่ไหนในการลงในใบ Manifest
    'ถ้า Customer Site ที่เลือกถูกระบุว่าเป็น isManifestAddr=Y ให้ใช้ที่อยู่นั้นเลย
    'แต่ถ้าไม่ได้ระบุให้ใช้ที่อยู่ทั่วไปของลูกค้า
    MsgBox "เสร็จสิ้นการพิมพ์ใบ Manifest แล้วจำนวน " & iCount & " รายการ"
End Sub

Private Sub Command2_Click()
 Dim x%, iCount%
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีรายการใบ Manifest ที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    If colControlDoc_ Is Nothing Then
        MsgBox "ไม่มีรายการใบคุมที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    If colControlDoc_.Count = 0 Then
        MsgBox "ไม่มีรายการใบคุมที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    
   
    
    
    
    PrintSelected = False
    For x = 1 To colControlDoc_.Count
        DBConnExc " UPDATE b  Set b.TruckRateID = c.TruckRateID FROM tbJobDataTimeTable a JOIN tbJobDataCarUsed b ON a.JobDataCarID = b.JobDataCarID  JOIN Com_vwUpdate_Manifest_Quotation_Minweight_Price c ON a.TimeTableNo = c.TimeTableNo  Where     (c.TruckRateID IS NOT NULL AND c.TruckRateID <> '')     AND (b.TruckRateID IS NULL OR b.TruckRateID = '')     AND c.WorkDate >= '2024-04-01' and a.TimeTableNo='" & colControlDoc_(x).TimeTableNo & "'"
        Set tmpCol = tmpS.JobDataTimeTableQuickSearch("TimeTableNo='" & colControlDoc_(x).TimeTableNo & "' And isMain='Y' And isCancel ='N'  AND ISNULL(isCanceled, 'N') <> 'Y' ")
        If tmpCol.Count > 0 Then
            iCount = iCount + 1
            PrintDocControl tmpCol(1)
        End If
        Set tmpCol = Nothing
    Next
    PrintSelected = False
    Set tmpS = Nothing
    If iCount = 0 Then
        MsgBox "กรุณาเลือกรายการใบคุมที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    MsgBox "เสร็จสิ้นการพิมพ์ใบคุมแล้วจำนวน " & iCount & " รายการ"
End Sub

Private Sub Command3_Click()
Dim x%
    For x = 1 To lvResult.ListItems.Count
        lvResult.ListItems(x).Checked = True
    Next
End Sub

Private Sub Command4_Click()
Dim x%
    For x = 1 To lvResult.ListItems.Count
        lvResult.ListItems(x).Checked = False
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim backDate As String, selX As Integer, x%
Dim tmpY As Integer, tmpM As Integer, tmpD As Integer
Dim tmpArr
Dim sqlStr As String

    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    selX = -1
    For x = 0 To Printers.Count - 1
        cboPrinter.AddItem Printers(x).DeviceName
        If Printers(x).DeviceName = Printer.DeviceName Then
            selX = x
        End If
    Next
    cboPrinter.ListIndex = selX
    sqlStr = "select backdate=convert(varchar(10),getdate()-7,111)"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    backDate = Trim("" & tmpRec("BackDate"))
    tmpRec.Close
    If IsDate(backDate) Then
        tmpArr = Split(backDate, "/")
        If CInt(tmpArr(0)) < 2500 Then
            dtFrom.ValueYMD = Trim("" & (CInt(tmpArr(0) + 543))) & "/" & Trim("" & tmpArr(1)) & "/" & Trim("" & tmpArr(2))
        Else
            dtFrom.ValueYMD = backDate
        End If
    End If
    Set tmpRec = Nothing
    With cboJobStatus
        .AddItem "== ทุกสถานะ =="
        .AddItem "งานใหม่"
        .AddItem "พิมพ์ Manifest แล้ว"
        .AddItem "ปิดงานแล้ว"
        .AddItem "ยกเลิก"
        .ListIndex = 0
        
        
    End With
    lvResult.ColumnHeaders(1).Icon = "down"
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
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
        Case 7, 8
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
        Case 10
            Select Case ColumnHeader.Icon
                Case "down"
                    ColumnHeader.Icon = "up"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortDescending, sortNumeric)
                Case "up"
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortNumeric)
                Case Else
                    ColumnHeader.Icon = "down"
                    Call SortColumn(lvResult, ColumnHeader.Index, sortAscending, sortNumeric)
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
Dim tmpCol As Collection
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    Set tmpCol = tmpS.MenifestSearch("MenifestID=" & lvResult.SelectedItem.Key)
    If tmpCol.Count > 0 Then
        Load frmMenifest
        frmMenifest.SetDataMenifest tmpCol(1)
        If lvResult.SelectedItem.SubItems(18) <> "0" Then
            frmMenifest.txtRate.Text = lvResult.SelectedItem.SubItems(18)
            Else
            frmMenifest.txtRate.Text = lvResult.SelectedItem.SubItems(17)
        End If
'        frmMenifest.SSTab1.TabVisible(2) = False
        frmMenifest.Show 1
    End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
End Sub

Private Sub lvResult_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim tmpItem As BWGJobDataTimeTable
    If colControlDoc_ Is Nothing Then Set colControlDoc_ = New Collection
    If item.Checked Then
        If isExist(colControlDoc_, "'" & item.SubItems(9) & "'") = False Then
            Set tmpItem = New BWGJobDataTimeTable
            tmpItem.TimeTableNo = item.SubItems(9)
            colControlDoc_.Add tmpItem, "'" & item.SubItems(9) & "'"
        End If
    Else
        If isExist(colControlDoc_, "'" & item.SubItems(9) & "'") Then
            colControlDoc_.Remove "'" & item.SubItems(9) & "'"
        End If
    End If
End Sub

Private Sub lvResult_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Pont As POINTAPI
    Dim ret As BTN_STYLE

    If Button = vbRightButton Then
        hMenu = CreatePopupMenu()
        AppendMenu hMenu, MF_STRING, 1, "เลือกทั้งหมด"
        AppendMenu hMenu, MF_SEPARATOR, -1, "SEP1"
        AppendMenu hMenu, MF_STRING, 2, "ไม่เลือกทั้งหมด"
'        AppendMenu hMenu, MF_STRING, 3, "ปิดงานใบเมนิเฟส"
'        AppendMenu hMenu, MF_STRING, 4, "ปิดงานใบเมนิเฟสที่เลือกทั้งหมด"
        'If LCase(CurrentUser.LoginName) = "admin" Then AppendMenu hMenu, MF_STRING, 4, "อนุมัติจ่าย"
        GetCursorPos Pont
        ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD, _
                                Pont.x, Pont.y, Me.hwnd, ByVal 0&)
        DestroyMenu hMenu

        If ret = 1 Then
            Command3_Click
        ElseIf ret = 2 Then
            Command4_Click
'        ElseIf ret = 3 Then
'            MNU_NotApprove_Click
'        ElseIf ret = 4 Then
'            MNU_SetApprove_Click
        End If
    End If
End Sub

Private Sub SearchQue()
'On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim tmpList As ListItem
Dim x As Long, y As Long
Dim sqlStr As String, strCond As String, StatusDesc As String
Dim tmpArr
Dim curColor
Dim subStr As String
Dim Rate As String
Dim CarRate As String
    
    Screen.MousePointer = 11
    Set colControlDoc_ = Nothing
    Set colControlDoc_ = New Collection
    sqlStr = "SELECT MenifestID,DocNo,MenifestNo,CustomerName,SubTypeDesc,TruckTypeDesc,TransportName1,TransportName2,WorkDate,CloseDate, CloseTime,isClosed,isCanceled,isPrinted,DriverName,JobNo,TimeTableNo,isTrailer,ApproveDate " & vbCrLf
    sqlStr = sqlStr & ", CloseStaff = (SELECT TOP(1) StaffFName + ' ' + StaffLName FROM tbCompanyStaff WHERE CloseStaffID = StaffID)" & vbCrLf
    sqlStr = sqlStr & ", ApproveStaffName, DriverTelNo,StaffFName,CustAumphurName,CustProvinceName,GPSSend,GPSRequest,SendDesc,RateSingle,RateTrailer" & vbCrLf
    sqlStr = sqlStr & ",Rate15Q,Rate24Q,RatePickup,carrate,rate10,JobDataCarID,NewCustJob,ratetrail "
    sqlStr = sqlStr & "From vw_MenifestQue "
    If Trim(txtJobNo.Text) <> "" Then
        strCond = " WHERE JobNo Like '%" & Trim(txtJobNo.Text) & "%' Or TimeTableNo Like '%" & Trim(txtJobNo.Text) & "%'"
    Else
        strCond = " WHERE (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "') "
        If Trim(txtKey.Text) <> "" And Replace(Trim(txtKey.Text), "*", "") <> "" Then
            strCond = strCond & " And (MenifestNo Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or DocNo Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or CustomerName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')"
        End If
        If Trim(txtKey2.Text) <> "" And Replace(Trim(txtKey2.Text), "*", "") <> "" Then
            strCond = strCond & " And (TransportName1 Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or TransportName2 Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or DriverName Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or CarRegisID Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%')"
        End If
        Select Case cboJobStatus.ListIndex
            Case 1 'งานใหม่
                subStr = " ((isPrinted<>'Y' Or isPrinted is null) and (isCanceled<>'Y' Or isCanceled is null) and (isClosed<>'Y' Or isClosed is null)) "
            Case 2 'พิมพ์ Menifest แล้ว
                subStr = " (isPrinted='Y' and (isCanceled<>'Y' Or isCanceled is null) and (isClosed<>'Y' Or isClosed is null)) "
            Case 3 'ปิดงาน
                subStr = " (isClosed='Y' and (ISNULL(isCanceled, 'N') <>'Y')) "
                'subStr = " (isClosed='Y' and (isCanceled <>'Y' Or isCanceled is null)) "
            Case 4 'ยกเลิก
                subStr = " (isCanceled='Y') "
        End Select
        If cboJobStatus.ListIndex > 0 Then
            If InStr(1, UCase(strCond), "WHERE") > 0 Then
                strCond = strCond & " And " & subStr
            Else
                strCond = strCond & " WHERE " & subStr
            End If
        End If
    End If
    curCondition_ = strCond

    sqlStr = sqlStr & strCond
    sqlStr = SelectWithCompanyID(sqlStr, False)
    
    If Trim(txtSale.Text) = "" Or Trim(txtSale.Text) = "*" Then Set SaleStaff_ = Nothing
    If Not SaleStaff_ Is Nothing Then
        sqlStr = sqlStr & vbCrLf & "AND UnderSaleStaffID = '" & SaleStaff_.id & "'" & vbCrLf
    End If
    
    sqlStr = sqlStr & " ORDER BY WorkDate, MenifestNo, DocNo"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    
    
    If tmpRec Is Nothing Then
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ"
        Exit Sub
    End If
    lvResult.ListItems.Clear
    Do While Not tmpRec.EOF
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpRec!menifestID & "'", "" & tmpRec!MenifestNo & ", " & Trim("" & tmpRec!DocNo))
        If Trim("" & tmpRec!isPrinted) <> "Y" And Trim("" & tmpRec!IsCanceled) <> "Y" And Trim("" & tmpRec!isClosed) <> "Y" Then
            curColor = lbOpen.ForeColor
            StatusDesc = lbOpen.Caption
        ElseIf Trim("" & tmpRec!isPrinted) = "Y" And Trim("" & tmpRec!IsCanceled) <> "Y" And Trim("" & tmpRec!isClosed) <> "Y" Then
            curColor = lbApproved.ForeColor
            StatusDesc = lbApproved.Caption
        ElseIf Trim("" & tmpRec!isClosed) = "Y" And Trim("" & tmpRec!IsCanceled) <> "Y" Then
            curColor = lbClosed.ForeColor
            StatusDesc = lbClosed.Caption
        ElseIf Trim("" & tmpRec!IsCanceled) = "Y" Then
            curColor = lbCancel.ForeColor
            StatusDesc = lbCancel.Caption
        End If
        tmpList.ForeColor = curColor
        tmpList.SubItems(1) = Trim("" & tmpRec!CustomerName)
        tmpList.SubItems(2) = Trim("" & tmpRec!JobNo)
        If Trim("" & tmpRec!isTrailer) = "Y" Then
            tmpList.SubItems(3) = Trim("" & tmpRec!TruckTypeDesc) & " (พ่วง), " & Trim("" & tmpRec!SubTypeDesc)
        Else
            tmpList.SubItems(3) = Trim("" & tmpRec!TruckTypeDesc) & ", " & Trim("" & tmpRec!SubTypeDesc)
        End If
        tmpList.SubItems(4) = Trim("" & tmpRec!TransportName2)
        tmpList.SubItems(5) = Trim("" & tmpRec!DriverName)
        tmpList.SubItems(6) = FormatYMD_to_DMY(Trim("" & tmpRec!WorkDate), "/", "/")
        tmpList.SubItems(7) = FormatYMD_to_DMY(Trim("" & tmpRec!ApproveDate), "/", "/")
        tmpList.SubItems(8) = StatusDesc
        tmpList.SubItems(9) = Trim("" & tmpRec!TimeTableNo)
        tmpList.SubItems(10) = Trim("" & tmpRec!CloseStaff)
        tmpList.SubItems(11) = Trim("" & tmpRec!ApproveStaffName)
        tmpList.SubItems(12) = Trim("" & tmpRec!DriverTelNo)
        tmpList.SubItems(13) = Trim("" & tmpRec!StaffFName)
        tmpList.SubItems(14) = Trim("" & tmpRec!CustAumphurName) & " " & Trim("" & tmpRec!CustProvinceName)
        If Trim("" & tmpRec!GPSSend) <> "1" And Trim("" & tmpRec!GPSRequest) = "Y" Then
            tmpList.SubItems(15) = "ยังไม่ได้ส่ง"
        ElseIf Trim("" & tmpRec!GPSSend) = "1" And Trim("" & tmpRec!GPSRequest) = "Y" Then
            tmpList.SubItems(15) = "ส่งแล้ว"
        Else
            tmpList.SubItems(15) = "-"
        End If
        If tmpRec!SendDesc <> "" Then
             tmpList.SubItems(16) = tmpRec!SendDesc
        End If
        
        If tmpRec!TruckTypeDesc = "รถ Roll off Truck - เดี่ยว" Or tmpRec!TruckTypeDesc = "รถ Lugger Box" Then
            If tmpRec!RateSingle <> "Null" Or tmpRec!RateSingle <> "" Then
                Rate = tmpRec!RateSingle
            Else
                Rate = "0"
            End If
        ElseIf tmpRec!TruckTypeDesc = "รถ Lugger Box - พ่วง" Or tmpRec!TruckTypeDesc = "รถ Roll off Truck - พ่วง" Then
            If tmpRec!RateTrailer <> "Null" Or tmpRec!RateTrailer <> "" Then
                Rate = tmpRec!RateTrailer
            Else
                Rate = "0"
            End If
         ElseIf tmpRec!TruckTypeDesc = "รถแท็งค์ 15 ลูกบาศก์เมตร" Then
            If tmpRec!Rate15Q <> "Null" Or tmpRec!Rate15Q <> "" Then
                Rate = tmpRec!Rate15Q
            Else
                Rate = 0
            End If
         ElseIf tmpRec!TruckTypeDesc = "รถแท็งค์ 24 ลูกบาศก์เมตร" Then
            If tmpRec!Rate24Q <> "Null" Or tmpRec!Rate24Q <> "" Then
                Rate = tmpRec!Rate24Q
            Else
                Rate = "0"
            End If
         ElseIf tmpRec!TruckTypeDesc = "รถแท็งค์ ISO 24 ลูกบาศก์เมตร" Then
            If tmpRec!Rate24Q <> "Null" Or tmpRec!Rate24Q <> "" Then
                Rate = tmpRec!Rate24Q
            Else
                Rate = "0"
            End If
         ElseIf tmpRec!TruckTypeDesc = "รถเทรลเลอร์" Then
            If tmpRec!RateTrail <> "Null" Or tmpRec!RateTrail <> "" Then
                Rate = tmpRec!RateTrail
            Else
                Rate = "0"
            End If
         ElseIf tmpRec!TruckTypeDesc = "รถกระบะ" Or tmpRec!TruckTypeDesc = "รถกระบะควบคุมอุณหภูมิ" Or tmpRec!TruckTypeDesc = "กระบะ" Then
            If tmpRec!RatePickup <> "Null" Or tmpRec!RatePickup <> "" Then
                Rate = tmpRec!RatePickup
            Else
                Rate = "0"
            End If
         ElseIf tmpRec!TruckTypeDesc = "รถสิบล้อ" Then
            If tmpRec!Rate10 <> "Null" Or tmpRec!Rate10 <> "" Then
                Rate = tmpRec!Rate10
            Else
                Rate = "0"
            End If
        Else
        Rate = "0"
        End If
        tmpList.SubItems(17) = Rate
        If tmpRec!CarRate <> "" Then
            CarRate = tmpRec!CarRate
            Else
            CarRate = "0"
        End If
        
        If CarRate = "0" Then
            DBConnExc "Update tbJobDataTimeTable Set CarRate='" & Rate & "' Where JobDataCarID='" & tmpRec!JobDataCarID & "'"
            tmpList.SubItems(18) = Rate
        Else
            tmpList.SubItems(18) = CarRate
        End If
        If tmpRec!NewCustJob <> "Null" Or tmpRec!NewCustJob <> "" Then
            tmpList.SubItems(19) = tmpRec!NewCustJob
        End If
        If tmpRec!NewCustJob <> "Null" Or tmpRec!NewCustJob <> "" Then
            tmpList.SubItems(19) = tmpRec!NewCustJob
        End If
        If tmpRec!JobDataCarID <> "Null" Or tmpRec!JobDataCarID <> "" Then
            tmpList.SubItems(20) = tmpRec!JobDataCarID
        End If
        
        For y = 1 To tmpList.ListSubItems.Count
            tmpList.ListSubItems(y).ForeColor = curColor
        Next
        Set tmpList = Nothing
        tmpRec.MoveNext
    Loop
    lvResult.ColumnHeaders(1).Icon = "down"
    Screen.MousePointer = 0
    frmMain.Stb1.Panels(1).Text = "พบข้อมูล " & Format(tmpRec.RecordCount, "#,##0") & " รายการ"
    tmpRec.Close
    Set tmpRec = Nothing
    Exit Sub
ErrD:
    Screen.MousePointer = 0
    GetMsg "", vbCritical
    Exit Sub
End Sub
Private Sub PrintTransControl(selItem As BWGJobDataTimeTable)
Dim tmpSql As String
Dim x%, y%, z%, iCount%, tmpCol As Collection
Dim tmpManifest As BWGMenifest, tmpStr As String, allCarRegisNo$
Dim UserGrp As Collection
Dim Group As New AWSUserGroup
Dim StrGrp As String
Dim i As Integer
        Dim myCRApp As New CRAXDRT.Application
        Dim myCRReport As CRAXDRT.Report
        Dim crxTable As CRAXDRT.DatabaseTable
    With rpt1
        Set tmpCol = selItem.GetMenifests("A")
        For x = 1 To tmpCol.Count
            Set tmpManifest = tmpCol(x)
            tmpStr = tmpManifest.DocNo
        .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
        .ReportFileName = App.Path & "\reports\TripMenifestDet.rpt"
        tmpSql = ""
        tmpSql = "SELECT * FROM vw_MenifestDetInTimeNo" & vbCrLf
        tmpSql = tmpSql & "WHERE docno= '" & tmpStr & "' " & vbCrLf
        .SQLQuery = tmpSql
        .Destination = crptToPrinter
        .PrintReport
        Next
        
    MsgBox "เสร็จสิ้นการพิมพ์ใบกำกับการขนส่งแล้ว"

        If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
    End With
End Sub
Private Sub PrintMain(selItem As BWGMenifest)
Dim pageCount As Integer
Dim maxRow As Integer, tmpDiscrep As BWGDiscrepWaste
Dim rowPerPage As Integer
Dim x%, y%, z%, iCount%
Dim YY As String
    rowPerPage = 12
    maxRow = selItem.JobDetails.Count
    pageCount = getPageCount(maxRow, rowPerPage)
    y = 1
    If IsDate(selItem.WorkDate) = True Then
        YY = Mid(selItem.WorkDate, 1, 4)
        YY = Right(YY, 2) & "/"
    Else
        YY = "N/A / "
    End If
            
    Set Printer = Printers(cboPrinter.ListIndex)
    
    Dim UserGrp As Collection
    Dim Group As New AWSUserGroup
    Dim StrGrp As String
    Dim i As Integer
        Set UserGrp = CurrentUser.getUserGroups
        For i = 1 To UserGrp.Count
            Set Group = UserGrp(i)
            If Group.Name = "Administrator" Then
                StrGrp = Group.Name
                Exit For
            Else
                StrGrp = Group.Name
            End If
        Next
        If StrGrp = "Administrator" Then
            If MsgBox("เลือก Export เป็นไฟล์ ตอบ Yes, เลือกพิมพ์เหมือนแผนกขนส่งตอบ No", vbInformation + vbYesNoCancel) = vbYes Then
                Dim myCRApp As New CRAXDRT.Application
                Dim myCRReport As CRAXDRT.Report
                Dim crxTable As CRAXDRT.DatabaseTable
                rpt1.ReportFileName = App.Path & "\reports\ManifestBill.rpt"
                Set myCRApp = New CRAXDRT.Application
                Set myCRReport = myCRApp.OpenReport(rpt1.ReportFileName)
                myCRReport.DiscardSavedData
                For Each crxTable In myCRReport.Database.Tables
                    crxTable.SetLogOnInfo DSN_NAME, strDb, strUid, strPwd
                Next
                
                Dim ExptoFile As Scripting.FileSystemObject
                Dim ExportFld As Folder
                Dim ExpReportPath As String
                Dim ManifestFile As String
                Set ExptoFile = New Scripting.FileSystemObject
''                ExpReportPath = "\\192.168.1.3\ftproot\ExportManifest\" & Format(selItem.WorkDate, "yyyyMMdd")
''                If ExptoFile.FolderExists(ExpReportPath) = False Then
''                    Set ExportFld = ExptoFile.CreateFolder(ExpReportPath)
''                End If
                                
                myCRReport.FormulaFields.GetItemByName("docNo").Text = Chr(34) & Replace(selItem.DocNo & "    " & selItem.CompanyName, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("ManifestNo").Text = Chr(34) & YY & Replace(selItem.MenifestNo, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("APManifestNo").Text = Chr(34) & Replace(selItem.CustomerCode & " " & selItem.APManifestNo, vbCrLf, "") & Chr(34)
                If selItem.CustTypeFlag = "C" Then
                    myCRReport.FormulaFields.GetItemByName("custName").Text = Chr(34) & Replace(selItem.CustomerName, vbCrLf, "") & Chr(34)
                ElseIf selItem.CustTypeFlag = "S" Then
                    myCRReport.FormulaFields.GetItemByName("custName").Text = Chr(34) & Replace(selItem.CustomerName & " (" & selItem.ParentCustName & ")", vbCrLf, "") & Chr(34)
                Else
                    myCRReport.FormulaFields.GetItemByName("custName").Text = Chr(34) & Replace(selItem.CustomerName, vbCrLf, "") & Chr(34)
                End If
                myCRReport.FormulaFields.GetItemByName("Item1").Text = Chr(34) & Replace(Replace(selItem.WasteGenNo, vbCrLf, ""), "DIW-G-", "") & Chr(34)
                If selItem.isManifestAddr Then
                    myCRReport.FormulaFields.GetItemByName("Item2").Text = Chr(34) & Replace(selItem.CustSiteAddr, vbCrLf, "") & Chr(34)
                Else
                    myCRReport.FormulaFields.GetItemByName("Item2").Text = Chr(34) & Replace(selItem.CustFullAddr, vbCrLf, "") & Chr(34)
                End If
                myCRReport.FormulaFields.GetItemByName("Item3").Text = Chr(34) & Replace(selItem.TelNo, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item4").Text = Chr(34) & Replace(selItem.FaxNo, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item5").Text = Chr(34) & Replace(selItem.ErContact, vbCrLf, "") & Chr(34)
                If selItem.CompanyID = "776EDA5830F744A6B3551348D851BC22" Then
                    myCRReport.FormulaFields.GetItemByName("Item6").Text = Chr(34) & Replace(selItem.TransportName1, vbCrLf, "") & Chr(34)
                Else
                    myCRReport.FormulaFields.GetItemByName("Item6").Text = Chr(34) & Replace(selItem.TransportName1, vbCrLf, "") & Chr(34)
                End If

                myCRReport.FormulaFields.GetItemByName("Item7").Text = Chr(34) & Replace(Replace(selItem.TransRegisNo1, "DIW-T-", ""), vbCrLf, "") & Chr(34)
                
                myCRReport.FormulaFields.GetItemByName("Item8").Text = Chr(34) & Replace(selItem.TransportName1, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item9").Text = Chr(34) & Replace(Replace(selItem.TransRegisNo1, "DIW-T-", ""), vbCrLf, "") & Chr(34)

                If InStr(Trim(selItem.TruckTypeName), "โรงงาน") = 0 Then
                    myCRReport.FormulaFields.GetItemByName("Item10").Text = Chr(34) & Replace(selItem.TSDFName1, vbCrLf, "") & Chr(34)
                    myCRReport.FormulaFields.GetItemByName("Item11").Text = Chr(34) & Replace(Replace(selItem.TSDFRegisNo1, "DIW-T-", ""), vbCrLf, "") & Chr(34)
                Else
                    myCRReport.FormulaFields.GetItemByName("Item10").Text = Chr(34) & "-" & Chr(34)
                    myCRReport.FormulaFields.GetItemByName("Item11").Text = Chr(34) & "-" & Chr(34)
                End If
                myCRReport.FormulaFields.GetItemByName("Item12").Text = Chr(34) & Replace(selItem.TSDFName2, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item13").Text = Chr(34) & Replace(selItem.TSDFRegisNo2, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item14").Text = Chr(34) & Replace(selItem.TotalLiquid, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item15").Text = Chr(34) & Replace(selItem.TotalSolid, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item16").Text = Chr(34) & Replace(selItem.SpecialComment, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item17").Text = Chr(34) & Replace(selItem.SignName1, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item18").Text = Chr(34) & "" & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item19").Text = Chr(34) & Right(selItem.SignDate1, 2) & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item20").Text = Chr(34) & Mid(selItem.SignDate1, 6, 2) & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item21").Text = Chr(34) & Left(selItem.SignDate1, 4) & Chr(34)
                If selItem.CompanyID = "3F764F1713974948A42CC0FAF926C730" Then
                    myCRReport.FormulaFields.GetItemByName("Item22").Text = Chr(34) & "" & Chr(34)
                Else
                    myCRReport.FormulaFields.GetItemByName("Item22").Text = Chr(34) & Replace(selItem.TransportName2, vbCrLf, "") & Chr(34)
                End If
                myCRReport.FormulaFields.GetItemByName("Item23").Text = Chr(34) & Replace(selItem.TruckTypeName & ", " & selItem.TruckSubTypeName, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item24").Text = Chr(34) & Replace(Replace(selItem.TransRegisNo2, "DIW-T-", ""), vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item25").Text = Chr(34) & Replace(selItem.TransTelNo1, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item26").Text = Chr(34) & Replace(selItem.TransFaxNo1, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item27").Text = Chr(34) & Replace(selItem.TransErContact1, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item28").Text = Chr(34) & Replace(selItem.CarRegisID, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item29").Text = Chr(34) & Replace(selItem.FromProvinceName, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item30").Text = Chr(34) & Replace(selItem.ToProvinceName, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item31").Text = Chr(34) & Replace(selItem.TimeSpending, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item32").Text = Chr(34) & Replace(selItem.SignName2, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item33").Text = Chr(34) & "" & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item34").Text = Chr(34) & Right(selItem.SignDate2, 2) & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item35").Text = Chr(34) & Mid(selItem.SignDate2, 6, 2) & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item36").Text = Chr(34) & Left(selItem.SignDate2, 4) & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item37").Text = Chr(34) & Replace(selItem.TSDFName1, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item38").Text = Chr(34) & Replace(selItem.TSDFRegisNo1, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item39").Text = Chr(34) & Replace(selItem.TSDFAddr & " " & selItem.TSDFTumbolName & " " & selItem.TSDFAumphurName & " " & selItem.TSDFProvinceName, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item40").Text = Chr(34) & Replace(selItem.TSDFTelNo, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item41").Text = Chr(34) & Replace(selItem.TSDFFaxNo, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item42").Text = Chr(34) & Replace(selItem.TSDFErContact, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item43").Text = Chr(34) & Replace(selItem.TreatmentPeriod, vbCrLf, "") & Chr(34)
                If Trim(selItem.PeriodUnit) = "D" Then
                    myCRReport.FormulaFields.GetItemByName("Item44").Text = Chr(34) & "P" & Chr(34)
                    myCRReport.FormulaFields.GetItemByName("Item45").Text = Chr(34) & "" & Chr(34)
                    myCRReport.FormulaFields.GetItemByName("Item46").Text = Chr(34) & "" & Chr(34)
                ElseIf Trim(selItem.PeriodUnit) = "M" Then
                    myCRReport.FormulaFields.GetItemByName("Item44").Text = Chr(34) & "" & Chr(34)
                    myCRReport.FormulaFields.GetItemByName("Item45").Text = Chr(34) & "P" & Chr(34)
                    myCRReport.FormulaFields.GetItemByName("Item46").Text = Chr(34) & "" & Chr(34)
                ElseIf Trim(selItem.PeriodUnit) = "Y" Then
                    myCRReport.FormulaFields.GetItemByName("Item44").Text = Chr(34) & "" & Chr(34)
                    myCRReport.FormulaFields.GetItemByName("Item45").Text = Chr(34) & "" & Chr(34)
                    myCRReport.FormulaFields.GetItemByName("Item46").Text = Chr(34) & "P" & Chr(34)
                End If
                myCRReport.FormulaFields.GetItemByName("Item47").Text = Chr(34) & Replace(selItem.ActualQty, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item48").Text = Chr(34) & Replace(selItem.SignName3, vbCrLf, "") & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item49").Text = Chr(34) & "" & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item50").Text = Chr(34) & Right(selItem.SignDate3, 2) & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item51").Text = Chr(34) & Mid(selItem.SignDate3, 6, 2) & Chr(34)
                myCRReport.FormulaFields.GetItemByName("Item52").Text = Chr(34) & Left(selItem.SignDate3, 4) & Chr(34)
                If selItem.DiscrepWastes.Count > 0 Then
                    Set tmpDiscrep = selItem.DiscrepWastes(1)
                    myCRReport.FormulaFields.GetItemByName("Item53").Text = Chr(34) & Replace(tmpDiscrep.WasteType, vbCrLf, "") & Chr(34)
                    myCRReport.FormulaFields.GetItemByName("Item54").Text = Chr(34) & Replace(tmpDiscrep.DiscrepQty, vbCrLf, "") & Chr(34)
                    If tmpDiscrep.ActionTaken = "A" Then
                        myCRReport.FormulaFields.GetItemByName("Item55").Text = Chr(34) & "P" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Item56").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Item57").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Item58").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Item59").Text = Chr(34) & "" & Chr(34)
                    ElseIf tmpDiscrep.ActionTaken = "B" Then
                        myCRReport.FormulaFields.GetItemByName("Item55").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Item56").Text = Chr(34) & "P" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Item57").Text = Chr(34) & Replace(tmpDiscrep.NewWasteNo, vbCrLf, "") & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Item58").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Item59").Text = Chr(34) & "" & Chr(34)
                    ElseIf tmpDiscrep.ActionTaken = "C" Then
                        myCRReport.FormulaFields.GetItemByName("Item55").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Item56").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Item57").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Item58").Text = Chr(34) & "P" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Item59").Text = Chr(34) & Replace(tmpDiscrep.ReasonAction, vbCrLf, "") & Chr(34)
                    End If
        '            .Formulas(62) = "Item60=" & Chr(34) & Right(tmpDiscrep.ReturnDate, 2) & Chr(34)
        '            .Formulas(63) = "Item61=" & Chr(34) & Mid(tmpDiscrep.ReturnDate, 6, 2) & Chr(34)
        '            .Formulas(64) = "Item62=" & Chr(34) & Right(tmpDiscrep.ReturnDate, 2) & Chr(34)
        '            .Formulas(65) = "Item63=" & Chr(34) & Replace(tmpDiscrep.ReturnMenifestNo, vbCrLf, "") & Chr(34)
        '            .Formulas(66) = "Item64=" & Chr(34) & Replace(tmpDiscrep.TSDFName, vbCrLf, "") & Chr(34)
        '            .Formulas(67) = "Item65=" & Chr(34) & "" & Chr(34)
                End If
                iCount = 1
                For x = 1 To pageCount
                    myCRReport.FormulaFields.GetItemByName("ShowPageNo").Text = Chr(34) & x & "/" & pageCount & Chr(34)
                    z = 1
                    Do Until y > maxRow Or y > x * rowPerPage
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_1").Text = Chr(34) & iCount & Chr(34)
                        'เพิ่ม Waste RDF
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_2").Text = Chr(34) & Replace(selItem.JobDetails(y).WasteName, vbCrLf, "") & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_3").Text = Chr(34) & selItem.JobDetails(y).WasteEUCode & " " & selItem.JobDetails(y).HAType & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_4").Text = Chr(34) & selItem.JobDetails(y).ContainerCount & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_5").Text = Chr(34) & selItem.JobDetails(y).ContainerTypeDesc & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_6").Text = Chr(34) & selItem.JobDetails(y).EstWasteQty & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_7").Text = Chr(34) & selItem.JobDetails(y).DisposerWeight & Chr(34)
                        iCount = iCount + 1
                        z = z + 1
                        y = y + 1
                    Loop
                    Do Until z > maxRow Or z > rowPerPage
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_1").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_2").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_3").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_4").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_5").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_6").Text = Chr(34) & "" & Chr(34)
                        myCRReport.FormulaFields.GetItemByName("Row" & z & "_7").Text = Chr(34) & "" & Chr(34)
                        z = z + 1
                    Loop
                Next
                myCRReport.ExportOptions.FormatType = crEFTCrystalReport
                myCRReport.ExportOptions.DestinationType = crEDTDiskFile
                ManifestFile = ExpReportPath & "\" & selItem.DocNo & "_" & Replace(selItem.CustomerName, Chr(13), "") & "_" & Replace(selItem.CarRegisID, ".", "") & ".rpt"
                If ExptoFile.FileExists(ManifestFile) = False Then
                    myCRReport.ExportOptions.DiskFileName = ManifestFile
                Else
                    ExptoFile.DeleteFile ManifestFile
                    myCRReport.ExportOptions.DiskFileName = ManifestFile
                End If
    
                myCRReport.Export False
                Set myCRApp = Nothing
                Set myCRReport = Nothing
                
    '                .Destination = crptToFile '(selItem.docNo & selItem.CustomerName)
    '                .PrintReport
            ElseIf vbNo Then
                GoTo F_No
            End If
        Else
F_No:
            If PrintSelected = False Then rpt1.PrinterSelect: PrintSelected = True
            With rpt1
                .LastErrorString = ""
                '.PrinterName = Printer.DeviceName
                '.PrinterDriver = Printer.DriverName
                '.PrinterPort = Printer.Port
                .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
            Dim tSql As String
            Dim rs As New ADODB.Recordset
                tSql = "SELECT * FROM tbJobDataHeader " & vbCrLf
                tSql = tSql & "where JobNo= '" & selItem.JobNo & "' and isjobs='D' " & vbCrLf
                Set rs = GetRS(tSql, adOpenForwardOnly, adLockOptimistic)
                If rs.RecordCount > 0 Then
                    .ReportFileName = App.Path & "\reports\MenifestRDF.rpt"
                    Dim tmpSql As String
                    tmpSql = "SELECT * FROM vw_MenifestDetailSearch " & vbCrLf
                    tmpSql = tmpSql & "WHERE  DocNo = '" & selItem.DocNo & "' order by wastename" & vbCrLf
                    
                    .Formulas(1) = "CusAddress=" & Chr(34) & Replace(selItem.CustFullAddr, vbCrLf, "") & Chr(34)
                    .Formulas(2) = "DriverName=" & Chr(34) & Replace(selItem.DriverName, vbCrLf, "") & Chr(34)
                    .Formulas(3) = "CaregisID=" & Chr(34) & Replace(selItem.CarRegisID, vbCrLf, "") & Chr(34)
                    .Formulas(4) = "CarType=" & Chr(34) & Replace(selItem.TruckTypeName, vbCrLf, "") & Chr(34)
                    .Formulas(5) = "TimetableNo=" & Chr(34) & Replace(selItem.TimeTableNo, vbCrLf, "") & Chr(34)
                    .Formulas(6) = "JobNum=" & Chr(34) & Replace(selItem.ActualQty, vbCrLf, "") & Chr(34)
                    .Formulas(7) = "WorkDate=" & Chr(34) & Replace(FormatYMD_to_DMY(selItem.WorkDate, "/", "/"), vbCrLf, "") & Chr(34)
                    
'                    .Formulas(7) = "WorkDate=" & Chr(34) & Replace(lvResult.SelectedItem.SubItems(6), vbCrLf, "") & Chr(34)

                    .SQLQuery = tmpSql
                     .Destination = crptToPrinter
                    .PrintReport
                    If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
                Else
                .ReportFileName = App.Path & "\reports\ManifestBill.rpt"
                .Formulas(0) = "DocNo=" & Chr(34) & Replace(selItem.DocNo & "  " & selItem.CompanyName, vbCrLf, "") & Chr(34)
                .Formulas(1) = "ManifestNo=" & Chr(34) & YY & Replace(selItem.MenifestNo, vbCrLf, "") & Chr(34)
                .Formulas(1000) = "APManifestNo=" & Chr(34) & Replace(selItem.APManifestNo, vbCrLf, "") & Chr(34)
                If selItem.CustTypeFlag = "C" Then
                    .Formulas(2) = "CustName=" & Chr(34) & Replace(selItem.CustomerName, vbCrLf, "") & Chr(34)
                ElseIf selItem.CustTypeFlag = "S" Then
                    .Formulas(2) = "CustName=" & Chr(34) & Replace(selItem.CustomerName & " (" & selItem.ParentCustName & ")", vbCrLf, "") & Chr(34)
                Else
                    .Formulas(2) = "CustName=" & Chr(34) & Replace(selItem.CustomerName, vbCrLf, "") & Chr(34)
                End If
                .Formulas(3) = "Item1=" & Chr(34) & Replace(Replace(selItem.WasteGenNo, vbCrLf, ""), "DIW-G-", "") & Chr(34)
                If selItem.isManifestAddr Then
                    .Formulas(4) = "Item2=" & Chr(34) & Replace(selItem.CustSiteAddr, vbCrLf, "") & Chr(34)
                Else
                    .Formulas(4) = "Item2=" & Chr(34) & Replace(selItem.CustFullAddr, vbCrLf, "") & Chr(34)
                End If
                .Formulas(5) = "Item3=" & Chr(34) & Replace(selItem.TelNo, vbCrLf, "") & Chr(34)
                .Formulas(6) = "Item4=" & Chr(34) & Replace(selItem.FaxNo, vbCrLf, "") & Chr(34)
                .Formulas(7) = "Item5=" & Chr(34) & Replace(selItem.ErContact, vbCrLf, "") & Chr(34)

                Dim CustTruck As String
                If InStr(Trim(selItem.TransportName2), "โรงงาน") > 0 Then
                    CustTruck = "รถโรงงานลูกค้า"
                Else
                    CustTruck = selItem.TransportName2
                End If
                If InStr(Trim(selItem.TruckTypeName), "โรงงาน") > 0 Then
                    .Formulas(8) = "Item6=" & Chr(34) & Replace(CustTruck, vbCrLf, "") & Chr(34)
                    .Formulas(10) = "Item8=" & Chr(34) & "" & Chr(34)
                    .Formulas(9) = "Item7=" & Chr(34) & Replace(Replace(selItem.TransRegisNo2, "DIW-T-", ""), vbCrLf, "") & Chr(34)
                    '.Formulas(9) = "Item7=" & Chr(34) & Replace(Replace(selItem.TransRegisNo1, "DIW-T-", ""), vbCrLf, "") & Chr(34)
                    .Formulas(11) = "Item9=" & Chr(34) & "" & Chr(34)
                Else
                    .Formulas(8) = "Item6=" & Chr(34) & Replace(selItem.TransportName1, vbCrLf, "") & Chr(34)
                    .Formulas(10) = "Item8=" & Chr(34) & Replace(CustTruck, vbCrLf, "") & Chr(34)
'                     .Formulas(1001) = "DriverTelNo=" & Chr(34) & IIf(Replace(selItem.DriverTelNo, vbCrLf, "") = "", "ไม่มีเบอร์คนขับ", Replace(selItem.DriverTelNo, vbCrLf, "")) & Chr(34)
                    .Formulas(9) = "Item7=" & Chr(34) & Replace(Replace(selItem.TransRegisNo1, "DIW-T-", ""), vbCrLf, "") & Chr(34)
                    .Formulas(11) = "Item9=" & Chr(34) & Replace(Replace(selItem.TransRegisNo2, "DIW-T-", ""), vbCrLf, "") & Chr(34)
                End If

                If InStr(Trim(selItem.TruckTypeName), "โรงงาน") = 0 Then
                    .Formulas(12) = "Item10=" & Chr(34) & Replace(selItem.TSDFName1, vbCrLf, "") & Chr(34)
                    .Formulas(13) = "Item11=" & Chr(34) & Replace(Replace(selItem.TSDFRegisNo1, "DIW-T-", ""), vbCrLf, "") & Chr(34)
                Else
                    .Formulas(12) = "Item10=" & Chr(34) & "-" & Chr(34)
                    .Formulas(13) = "Item11=" & Chr(34) & "-" & Chr(34)
                End If
                .Formulas(14) = "Item12=" & Chr(34) & Replace(selItem.TSDFName2, vbCrLf, "") & Chr(34)
                .Formulas(15) = "Item13=" & Chr(34) & Replace(selItem.TSDFRegisNo2, vbCrLf, "") & Chr(34)
                .Formulas(16) = "Item14=" & Chr(34) & Replace(selItem.TotalLiquid, vbCrLf, "") & Chr(34)
                .Formulas(17) = "Item15=" & Chr(34) & Replace(selItem.TotalSolid, vbCrLf, "") & Chr(34)
                .Formulas(18) = "Item16=" & Chr(34) & Replace(selItem.SpecialComment, vbCrLf, "") & Chr(34)
                .Formulas(19) = "Item17=" & Chr(34) & Replace(selItem.SignName1, vbCrLf, "") & Chr(34)
                .Formulas(20) = "Item18=" & Chr(34) & "" & Chr(34)
                .Formulas(21) = "Item19=" & Chr(34) & Right(selItem.SignDate1, 2) & Chr(34)
                .Formulas(22) = "Item20=" & Chr(34) & Mid(selItem.SignDate1, 6, 2) & Chr(34)
                .Formulas(23) = "Item21=" & Chr(34) & Left(selItem.SignDate1, 4) & Chr(34)
                .Formulas(24) = "Item22=" & Chr(34) & Replace(selItem.TransportName2, vbCrLf, "") & Chr(34)
                .Formulas(25) = "Item23=" & Chr(34) & Replace(selItem.TruckTypeName & ", " & selItem.TruckSubTypeName, vbCrLf, "") & Chr(34)
                .Formulas(26) = "Item24=" & Chr(34) & Replace(Replace(selItem.TransRegisNo2, "DIW-T-", ""), vbCrLf, "") & Chr(34)
                .Formulas(27) = "Item25=" & Chr(34) & Replace(selItem.TransTelNo1, vbCrLf, "") & Chr(34)
                .Formulas(28) = "Item26=" & Chr(34) & Replace(selItem.TransFaxNo1, vbCrLf, "") & Chr(34)
                .Formulas(29) = "Item27=" & Chr(34) & Replace(selItem.TransErContact1, vbCrLf, "") & Chr(34)
                .Formulas(30) = "Item28=" & Chr(34) & Replace(selItem.CarRegisID, vbCrLf, "") & Chr(34)
                .Formulas(31) = "Item29=" & Chr(34) & Replace(selItem.FromProvinceName, vbCrLf, "") & Chr(34)
                .Formulas(32) = "Item30=" & Chr(34) & Replace(selItem.ToProvinceName, vbCrLf, "") & Chr(34)
                .Formulas(33) = "Item31=" & Chr(34) & Replace(selItem.TimeSpending, vbCrLf, "") & Chr(34)
                .Formulas(34) = "Item32=" & Chr(34) & Replace(selItem.SignName2, vbCrLf, "") & Chr(34)
                .Formulas(35) = "Item33=" & Chr(34) & "" & Chr(34)
                .Formulas(36) = "Item34=" & Chr(34) & Right(selItem.SignDate2, 2) & Chr(34)
                .Formulas(37) = "Item35=" & Chr(34) & Mid(selItem.SignDate2, 6, 2) & Chr(34)
                .Formulas(38) = "Item36=" & Chr(34) & Left(selItem.SignDate2, 4) & Chr(34)
                .Formulas(39) = "Item37=" & Chr(34) & Replace(selItem.TSDFName1, vbCrLf, "") & Chr(34)
                .Formulas(40) = "Item38=" & Chr(34) & Replace(selItem.TSDFRegisNo1, vbCrLf, "") & Chr(34)
                .Formulas(41) = "Item39=" & Chr(34) & Replace(selItem.TSDFAddr & " " & selItem.TSDFTumbolName & " " & selItem.TSDFAumphurName & " " & selItem.TSDFProvinceName, vbCrLf, "") & Chr(34)
                .Formulas(42) = "Item40=" & Chr(34) & Replace(selItem.TSDFTelNo, vbCrLf, "") & Chr(34)
                .Formulas(43) = "Item41=" & Chr(34) & Replace(selItem.TSDFFaxNo, vbCrLf, "") & Chr(34)
                .Formulas(44) = "Item42=" & Chr(34) & Replace(selItem.TSDFErContact, vbCrLf, "") & Chr(34)
                .Formulas(45) = "Item43=" & Chr(34) & Replace(selItem.TreatmentPeriod, vbCrLf, "") & Chr(34)
                If Trim(selItem.PeriodUnit) = "D" Then
                    .Formulas(46) = "Item44=" & Chr(34) & "P" & Chr(34)
                    .Formulas(47) = "Item45=" & Chr(34) & "" & Chr(34)
                    .Formulas(48) = "Item46=" & Chr(34) & "" & Chr(34)
                ElseIf Trim(selItem.PeriodUnit) = "M" Then
                    .Formulas(46) = "Item44=" & Chr(34) & "" & Chr(34)
                    .Formulas(47) = "Item45=" & Chr(34) & "P" & Chr(34)
                    .Formulas(48) = "Item46=" & Chr(34) & "" & Chr(34)
                ElseIf Trim(selItem.PeriodUnit) = "Y" Then
                    .Formulas(46) = "Item44=" & Chr(34) & "" & Chr(34)
                    .Formulas(47) = "Item45=" & Chr(34) & "" & Chr(34)
                    .Formulas(48) = "Item46=" & Chr(34) & "P" & Chr(34)
                End If
                .Formulas(49) = "Item47=" & Chr(34) & Replace(selItem.ActualQty, vbCrLf, "") & Chr(34)
                .Formulas(50) = "Item48=" & Chr(34) & Replace(selItem.SignName3, vbCrLf, "") & Chr(34)
                .Formulas(51) = "Item49=" & Chr(34) & "" & Chr(34)
                .Formulas(52) = "Item50=" & Chr(34) & Right(selItem.SignDate3, 2) & Chr(34)
                .Formulas(53) = "Item51=" & Chr(34) & Mid(selItem.SignDate3, 6, 2) & Chr(34)
                .Formulas(54) = "Item52=" & Chr(34) & Left(selItem.SignDate3, 4) & Chr(34)
                If selItem.DiscrepWastes.Count > 0 Then
                    Set tmpDiscrep = selItem.DiscrepWastes(1)
                    .Formulas(55) = "Item53=" & Chr(34) & Replace(tmpDiscrep.WasteType, vbCrLf, "") & Chr(34)
                    .Formulas(56) = "Item54=" & Chr(34) & Replace(tmpDiscrep.DiscrepQty, vbCrLf, "") & Chr(34)
                    If tmpDiscrep.ActionTaken = "A" Then
                        .Formulas(57) = "Item55=" & Chr(34) & "P" & Chr(34)
                        .Formulas(58) = "Item56=" & Chr(34) & "" & Chr(34)
                        .Formulas(59) = "Item57=" & Chr(34) & "" & Chr(34)
                        .Formulas(60) = "Item58=" & Chr(34) & "" & Chr(34)
                        .Formulas(61) = "Item59=" & Chr(34) & "" & Chr(34)
                    ElseIf tmpDiscrep.ActionTaken = "B" Then
                        .Formulas(57) = "Item55=" & Chr(34) & "" & Chr(34)
                        .Formulas(58) = "Item56=" & Chr(34) & "P" & Chr(34)
                        .Formulas(59) = "Item57=" & Chr(34) & Replace(tmpDiscrep.NewWasteNo, vbCrLf, "") & Chr(34)
                        .Formulas(60) = "Item58=" & Chr(34) & "" & Chr(34)
                        .Formulas(61) = "Item59=" & Chr(34) & "" & Chr(34)
                    ElseIf tmpDiscrep.ActionTaken = "C" Then
                        .Formulas(57) = "Item55=" & Chr(34) & "" & Chr(34)
                        .Formulas(58) = "Item56=" & Chr(34) & "" & Chr(34)
                        .Formulas(59) = "Item57=" & Chr(34) & "" & Chr(34)
                        .Formulas(60) = "Item58=" & Chr(34) & "P" & Chr(34)
                        .Formulas(61) = "Item59=" & Chr(34) & Replace(tmpDiscrep.ReasonAction, vbCrLf, "") & Chr(34)
                    End If
                    .Formulas(62) = "Item60=" & Chr(34) & Right(tmpDiscrep.ReturnDate, 2) & Chr(34)
                    .Formulas(63) = "Item61=" & Chr(34) & Mid(tmpDiscrep.ReturnDate, 6, 2) & Chr(34)
                    .Formulas(64) = "Item62=" & Chr(34) & Right(tmpDiscrep.ReturnDate, 2) & Chr(34)
                    .Formulas(65) = "Item63=" & Chr(34) & Replace(tmpDiscrep.ReturnMenifestNo, vbCrLf, "") & Chr(34)
                    .Formulas(66) = "Item64=" & Chr(34) & Replace(tmpDiscrep.TSDFName, vbCrLf, "") & Chr(34)
                    .Formulas(67) = "Item65=" & Chr(34) & "" & Chr(34)
                End If
                iCount = 1
                For x = 1 To pageCount
                    .Formulas(68) = "ShowPageNo=" & Chr(34) & x & "/" & pageCount & Chr(34)
                    z = 1
                    Do Until y > maxRow Or y > x * rowPerPage
                        .Formulas(12 * (z - 1) + 69) = "Row" & z & "_1=" & Chr(34) & iCount & Chr(34)
                        'เพิ่มประเภท Waste RDF
                        .Formulas(12 * (z - 1) + 70) = "Row" & z & "_2=" & Chr(34) & Replace(selItem.JobDetails(y).WasteName, vbCrLf, "") & Chr(34)
                        .Formulas(12 * (z - 1) + 71) = "Row" & z & "_3=" & Chr(34) & selItem.JobDetails(y).WasteEUCode & " " & selItem.JobDetails(y).HAType & Chr(34)
                        .Formulas(12 * (z - 1) + 72) = "Row" & z & "_4=" & Chr(34) & selItem.JobDetails(y).ContainerCount & Chr(34)
                        .Formulas(12 * (z - 1) + 73) = "Row" & z & "_5=" & Chr(34) & selItem.JobDetails(y).ContainerTypeDesc & Chr(34)
                        .Formulas(12 * (z - 1) + 74) = "Row" & z & "_6=" & Chr(34) & selItem.JobDetails(y).EstWasteQty & Chr(34)
                        .Formulas(12 * (z - 1) + 75) = "Row" & z & "_7=" & Chr(34) & selItem.JobDetails(y).DisposerWeight & Chr(34)
                        iCount = iCount + 1
                        z = z + 1
                        y = y + 1
                    Loop
                    .LastErrorString = ""
                    Do Until z > maxRow Or z > rowPerPage
                        .Formulas(12 * (z - 1) + 69) = "Row" & z & "_1=" & Chr(34) & "" & Chr(34)
                        .Formulas(12 * (z - 1) + 70) = "Row" & z & "_2=" & Chr(34) & "" & Chr(34)
                        .Formulas(12 * (z - 1) + 71) = "Row" & z & "_3=" & Chr(34) & "" & Chr(34)
                        .Formulas(12 * (z - 1) + 72) = "Row" & z & "_4=" & Chr(34) & "" & Chr(34)
                        .Formulas(12 * (z - 1) + 73) = "Row" & z & "_5=" & Chr(34) & "" & Chr(34)
                        .Formulas(12 * (z - 1) + 74) = "Row" & z & "_6=" & Chr(34) & "" & Chr(34)
                        .Formulas(12 * (z - 1) + 75) = "Row" & z & "_7=" & Chr(34) & "" & Chr(34)
                        z = z + 1
                    Loop
                    If Trim(.LastErrorString) <> "" Then
                        MsgBox .LastErrorString
                        .LastErrorString = ""
                    End If
        
                    .Destination = crptToPrinter
                    .PrintReport
                    If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
                Next
                '.Reset
                .Formulas(0) = "DocNo=" & Chr(34) & "" & Chr(34)
                .Formulas(1) = "ManifestNo=" & Chr(34) & "" & Chr(34)
                .Formulas(2) = "CustName=" & Chr(34) & "" & Chr(34)
                For x = 3 To 67
                    .Formulas(x) = "Item" & (x - 2) & "=" & Chr(34) & "" & Chr(34)
                Next
                .Formulas(68) = "ShowPageNo=" & Chr(34) & "" & Chr(34)
                For x = 1 To 12
                    .Formulas(12 * (x - 1) + 69) = "Row" & x & "_1=" & Chr(34) & "" & Chr(34)
                    .Formulas(12 * (x - 1) + 70) = "Row" & x & "_2=" & Chr(34) & "" & Chr(34)
                    .Formulas(12 * (x - 1) + 71) = "Row" & x & "_3=" & Chr(34) & "" & Chr(34)
                    .Formulas(12 * (x - 1) + 72) = "Row" & x & "_4=" & Chr(34) & "" & Chr(34)
                    .Formulas(12 * (x - 1) + 73) = "Row" & x & "_5=" & Chr(34) & "" & Chr(34)
                    .Formulas(12 * (x - 1) + 74) = "Row" & x & "_6=" & Chr(34) & "" & Chr(34)
                    .Formulas(12 * (x - 1) + 75) = "Row" & x & "_7=" & Chr(34) & "" & Chr(34)
                Next
                 End If
                If (StrGrp <> "Administrator") Then
                    DBConnExc "Update tbMenifestHeader Set isPrinted='Y' WHERE MenifestID='" & selItem.id & "'"
                End If
            End With
        End If
End Sub

Private Function getPageCount(maxRow As Integer, rowByPage As Integer) As Integer
Dim num1%, num2%
Dim pageCount As Integer
Dim tmpStr As String
Dim rowCount As Long
    num1 = maxRow
    num2 = rowByPage
    rowCount = 0
    If num1 > num2 Then
        Do Until rowCount >= maxRow
            pageCount = pageCount + 1
            rowCount = rowByPage * pageCount
        Loop
        getPageCount = pageCount
    Else
        getPageCount = 1
    End If
End Function

Private Sub PrintDocControl(selItem As BWGJobDataTimeTable)
Dim x%, y%, z%, iCount%, tmpCol As Collection
Dim tmpManifest As BWGMenifest, tmpStr As String, allCarRegisNo$
Dim UserGrp As Collection
Dim Group As New AWSUserGroup
Dim StrGrp As String
Dim i As Integer
Dim con As String
    Set UserGrp = CurrentUser.getUserGroups
'    Set Printer = Printers(cboPrinter.ListIndex)
        '************
    For i = 1 To UserGrp.Count
        Set Group = UserGrp(i)
        If Group.Name = "Administrator" Then
            StrGrp = Group.Name
            Exit For
        Else
            StrGrp = Group.Name
        End If
    Next
     Dim tmpWaste As ADODB.Recordset
        Set tmpWaste = GetRS("select b.CustomerCode,b.WasteName from vw_MenifestSearch a join vw_MenifestDetail b on a.MenifestID=b.MenifestID where TimeTableNo='" & selItem.TimeTableNo & "'", adOpenForwardOnly, adLockReadOnly)
    If StrGrp = "Administrator" Then
        Dim myCRApp As New CRAXDRT.Application
        Dim myCRReport As CRAXDRT.Report
        Dim crxTable As CRAXDRT.DatabaseTable
        rpt1.ReportFileName = App.Path & "\reports\TripControl.rpt"
        Set myCRApp = New CRAXDRT.Application
        Set myCRReport = myCRApp.OpenReport(rpt1.ReportFileName)
        myCRReport.DiscardSavedData
        For Each crxTable In myCRReport.Database.Tables
            crxTable.SetLogOnInfo DSN_NAME, strDb, strUid, strPwd
        Next
        
        Dim ExptoFile As Scripting.FileSystemObject
        Dim ExportFld As Folder
        Dim ExpReportPath As String
        Dim TripControlFile As String
        Dim WasteCount As Integer
        Set ExptoFile = New Scripting.FileSystemObject
        ExpReportPath = "\\192.168.1.2\ftproot\ExportManifest\" & Format(selItem.WorkDate, "yyyyMMdd")
        If ExptoFile.FolderExists(ExpReportPath) = False Then
            Set ExportFld = ExptoFile.CreateFolder(ExpReportPath)
        End If
        
        
       
    
    
        myCRReport.FormulaFields.GetItemByName("ShowDay").Text = Chr(34) & Right(selItem.WorkDate, 2) & Chr(34)
        myCRReport.FormulaFields.GetItemByName("ShowMonth").Text = Chr(34) & Mid(selItem.WorkDate, 6, 2) & Chr(34)
        myCRReport.FormulaFields.GetItemByName("ShowYear").Text = Chr(34) & Left(selItem.WorkDate, 4) & Chr(34)
        Set tmpCol = selItem.GetMenifests("A")
        If tmpCol.Count > 0 Then
            Set tmpManifest = tmpCol(1)
        Else
            Set tmpManifest = New BWGMenifest
        End If
        Dim PhoDesW, PhoDesB, PhoDesBF, CustPhoto, CustCon, EmpTrain, CarCheck, CloseWeight, NewCustJob As String
        Dim Photo As String
        
            If tmpManifest.TakePhoto = "Y" Then
                Photo = "ถ่ายรูปทำลาย"
            ElseIf tmpManifest.TakePhoto = "W" Then
                Photo = "ถ่ายรูปชั่ง + ลงของ"
                
            End If

'        If tmpManifest.PhotoOnWeight = "Y" Then
'            PhoDesW = "ถ่ายรูปชั่ง + ลงของ"
'        End If
'        If tmpManifest.PhotoOnBelt = "Y" Then
'            PhoDesB = "ถ่ายรูปทำลาย"
'        End If
        If tmpManifest.PhotoOnBoxfeed = "Y" Then
            PhoDesBF = "ถ่ายรูป Boxfeed"
        End If
         If tmpManifest.CustCon = "Y" Then
            CustCon = "เก็บภาชนะไว้คืนลูกค้า"
        End If
         If tmpManifest.CustPhoto = "Y" Then
            CustPhoto = "ลูกค้าตามดูทำลาย"
        End If
         If tmpManifest.EmpTrain = "Y" Then
            EmpTrain = "อบรมพนักงาน"
        End If
         If tmpManifest.CarCheck = "Y" Then
            CarCheck = "ตรวจสภาพรถ"
        End If
         If tmpManifest.NewCustJob = "Y" Then
            NewCustJob = "งานลูกค้าใหม่"
        End If
       
        myCRReport.FormulaFields.GetItemByName("CompanyName").Text = Chr(34) & Replace(tmpManifest.CustomerName, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("Vendor").Text = Chr(34) & Replace(tmpManifest.TransportName2, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("ShowAddr").Text = Chr(34) & Replace(tmpManifest.CustSiteAddr, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("ContactPerson").Text = Chr(34) & Replace(tmpManifest.ErContact, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("CarType").Text = Chr(34) & Replace(tmpManifest.TruckTypeName & ", " & tmpManifest.TruckSubTypeName, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("TelNo").Text = Chr(34) & Replace(tmpManifest.TelNo, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("DriverName").Text = Chr(34) & Replace(tmpManifest.DriverName, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("AppointDate").Text = Chr(34) & Replace(FormatYMD_to_DMY(selItem.AppointDate, "/", "/"), vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("WorkerCount").Text = Chr(34) & Replace(selItem.WorkerCount, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("TimeBack").Text = Chr(34) & Replace(tmpManifest.OilUse, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("TimeIn").Text = Chr(34) & Replace(tmpManifest.TimeArriveCust, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("TimeIn").Text = Chr(34) & Replace(tmpManifest.StartTime, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("TimeOut").Text = Chr(34) & Replace(tmpManifest.TimeOutCust, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("SaleStaff").Text = Chr(34) & Replace("เจ้าหน้าที่การตลาด คุณ" & tmpManifest.UnderStaffName & "   เบอร์โทร. " & selItem.StTel, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("Remmark").Text = Chr(34) & Replace(tmpManifest.CarRem, vbCrLf, "") & Chr(34)
        WasteCount = 1
         Do Until tmpWaste.EOF
            myCRReport.FormulaFields.GetItemByName("WasteName" & WasteCount).Text = Chr(34) & CStr(WasteCount) & ". " & Trim(tmpWaste!CustomerCode) & " " & Trim(tmpWaste!WasteName) & Chr(34)
            tmpWaste.MoveNext
            WasteCount = WasteCount + 1
            If WasteCount >= 10 Then
                 tmpWaste.MoveLast
                 tmpWaste.MoveNext
            End If
        Loop

        
        
        If tmpManifest.CarRate <> "0" Then
        myCRReport.FormulaFields.GetItemByName("Rate").Text = Chr(34) & Replace("เรทน้ำมัน " & tmpManifest.CarRate & " ลิตร", vbCrLf, "") & Chr(34)
        Else
        myCRReport.FormulaFields.GetItemByName("Rate").Text = ""
        End If
        myCRReport.FormulaFields.GetItemByName("RemPhoto").Text = Chr(34) & Replace(Photo & " " & PhoDesW & " " & PhoDesB & " " & PhoDesBF & " " & CustPhoto & " " & CustCon & " " & CloseWeight & " " & EmpTrain & " " & CarCheck & " " & NewCustJob, vbCrLf, "") & Chr(34)
        con = ""
        If tmpManifest.Tankcount = "" Then tmpManifest.Tankcount = "0"
        If tmpManifest.Tankcount > 0 Then
        con = "ถังปากกว้าง " & tmpManifest.Tankcount & " "
        End If
        If tmpManifest.TankTypeSCount = "" Then tmpManifest.TankTypeSCount = "0"
        If tmpManifest.TankTypeSCount > 0 Then
        con = con & "ถังปากจู๋ " & tmpManifest.TankTypeSCount & " "
        End If
        If tmpManifest.BagCount = "" Then tmpManifest.BagCount = "0"
        If tmpManifest.BagCount > 0 Then
        con = con & "ถุงเปิดปาก " & tmpManifest.BagCount & " "
        End If
        If tmpManifest.BagBack = "" Then tmpManifest.BagBack = "0"
        If tmpManifest.BagBack > 0 Then
        con = con & "ถุงดำ " & tmpManifest.BagBack & " "
        End If
        If tmpManifest.Container = "" Then tmpManifest.Container = "0"
        If tmpManifest.Container > 0 Then
        con = con & "เบ้าท์ใส่น้ำเสีย " & tmpManifest.Container & " "
        End If
        If tmpManifest.Lack = "" Then tmpManifest.Lack = "0"
        If tmpManifest.Lack > 0 Then
        con = con & "แลค " & tmpManifest.Lack & " "
        End If
        If tmpManifest.PaperBox = "" Then tmpManifest.PaperBox = "0"
        If tmpManifest.PaperBox > 0 Then
        con = con & "กล่องกระดาษ " & tmpManifest.PaperBox & " "
        End If
        If tmpManifest.BoutS = "" Then tmpManifest.BoutS = "0"
        If tmpManifest.BoutS > 0 Then
        con = con & "เบ้าท์ตัดปาก " & tmpManifest.BoutS & " "
        End If
        If tmpManifest.LackNew = "" Then tmpManifest.LackNew = "0"
        If tmpManifest.LackNew > 0 Then
        con = con & "แลครุ่นใหม่ " & tmpManifest.LackNew & " "
        End If
        If tmpManifest.LackGas = "" Then tmpManifest.LackGas = "0"
         If tmpManifest.LackGas > 0 Then
        con = con & "แลคใส่ถังแก๊ส " & tmpManifest.LackGas & " "
        End If
        If tmpManifest.Plate = "" Then tmpManifest.Plate = "0"
         If tmpManifest.Plate > 0 Then
        con = con & "พาเลท " & tmpManifest.Plate & " "
        End If
        If tmpManifest.Plastic = "" Then tmpManifest.Plastic = "0"
         If tmpManifest.Plastic > 0 Then
        con = con & "พลาสติกแล็ป " & tmpManifest.Plastic & " "
        End If
        If tmpManifest.BoutCage = "" Then tmpManifest.BoutCage = "0"
         If tmpManifest.BoutCage > 0 Then
        con = con & "กรงเบ้าท์ " & tmpManifest.BoutCage & ""
        End If
        If tmpManifest.CovidBin = "" Then tmpManifest.CovidBin = "0"
        If tmpManifest.CovidBin > 0 Then
        con = con & "ถังขยะโควิด " & tmpManifest.CovidBin & " "
        End If
       
        myCRReport.FormulaFields.GetItemByName("Con").Text = Chr(34) & con & Chr(34)
        
        For x = 1 To tmpCol.Count
            Set tmpManifest = tmpCol(x)
            tmpStr = tmpStr & tmpManifest.DocNo & ", "
            If InStr(1, allCarRegisNo, tmpManifest.CarRegisID) = 0 Then
                allCarRegisNo = allCarRegisNo & tmpManifest.CarRegisID & ", "
            End If
        Next
        allCarRegisNo = Left(Trim(allCarRegisNo), Len(Trim(allCarRegisNo)) - 1)
        myCRReport.FormulaFields.GetItemByName("CarRegisID").Text = Chr(34) & Replace(allCarRegisNo, vbCrLf, "") & Chr(34)
        
        If Trim(tmpStr) <> "" Then
            tmpStr = Left(Trim(tmpStr), Len(Trim(tmpStr)) - 1)
            myCRReport.FormulaFields.GetItemByName("MenifestList").Text = Chr(34) & selItem.TimeTableNo & " : " & tmpStr & Chr(34)
        End If
    
        myCRReport.ExportOptions.FormatType = crEFTCrystalReport
        myCRReport.ExportOptions.DestinationType = crEDTDiskFile
        TripControlFile = ExpReportPath & "\" & selItem.TimeTableNo & "_" & Replace(tmpManifest.CustomerName & "_" & allCarRegisNo, Chr(13), "") & ".rpt"
        If ExptoFile.FileExists(TripControlFile) = False Then
            myCRReport.ExportOptions.DiskFileName = TripControlFile
        Else
            ExptoFile.DeleteFile TripControlFile
            myCRReport.ExportOptions.DiskFileName = TripControlFile
        End If
        
        myCRReport.Export False
        Set myCRApp = Nothing
        Set myCRReport = Nothing
        Set tmpManifest = Nothing
    Else
        If PrintSelected = False Then rpt2.PrinterSelect: PrintSelected = True
        With rpt2
    '        .PrinterName = Printer.DeviceName
    '        .PrinterDriver = Printer.DriverName
    '        .PrinterPort = Printer.Port
            .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
            .ReportFileName = App.Path & "\reports\TripControl.rpt"
            
            .Formulas(0) = "ShowDay=" & Chr(34) & Right(selItem.WorkDate, 2) & Chr(34)
            .Formulas(1) = "ShowMonth=" & Chr(34) & Mid(selItem.WorkDate, 6, 2) & Chr(34)
            .Formulas(2) = "ShowYear=" & Chr(34) & Left(selItem.WorkDate, 4) & Chr(34)
            Set tmpCol = selItem.GetMenifests("A")
            If tmpCol.Count > 0 Then
                Set tmpManifest = tmpCol(1)
            Else
                Set tmpManifest = New BWGMenifest
            End If
            If tmpManifest.TakePhoto = "Y" Then
                Photo = "ถ่ายรูปทำลาย"
            'ElseIf tmpManifest.TakePhoto = "W" Then
                'Photo = "ถ่ายรูปชั่ง + ลงของ"
            ElseIf tmpManifest.TakePhoto = "S" Then
                Photo = "ถ่ายรูปพิเศษ"
'            ElseIf tmpManifest.TakePhoto = "N" Then
'                Photo = "ถ่ายรูปชั่ง + ลงของ"
            End If
'            If tmpManifest.PhotoOnWeight = "Y" Then
'                PhoDesW = "ถ่ายรูปชั่ง + ลงของ"
'            End If
'            If tmpManifest.PhotoOnBelt = "Y" Then
'                PhoDesB = "ถ่ายรูปทำลาย"
'            End If
            If tmpManifest.PhotoOnBoxfeed = "Y" Then
                PhoDesBF = "ถ่ายรูป Boxfeed"
            End If
             If tmpManifest.CustCon = "Y" Then
                CustCon = "เก็บภาชนะไว้คืนลูกค้า"
            End If
             If tmpManifest.CustPhoto = "Y" Then
                CustPhoto = "ลูกค้าตามดูทำลาย"
            End If
             If tmpManifest.EmpTrain = "Y" Then
                EmpTrain = "อบรมพนักงาน"
            End If
             If tmpManifest.CarCheck = "Y" Then
                CarCheck = "ตรวจสภาพรถ"
            End If
             If tmpManifest.NewCustJob = "Y" Then
                NewCustJob = "งานลูกค้าใหม่"
            End If

            .Formulas(3) = "CompanyName=" & Chr(34) & Replace(tmpManifest.CustomerName, vbCrLf, "") & Chr(34)
            .Formulas(4) = "Vendor=" & Chr(34) & Replace(tmpManifest.TransportName2, vbCrLf, "") & Chr(34)
            .Formulas(5) = "ShowAddr=" & Chr(34) & Replace(tmpManifest.CustSiteAddr, vbCrLf, "") & Chr(34)
            .Formulas(7) = "ContactPerson=" & Chr(34) & Replace(tmpManifest.ErContact, vbCrLf, "") & Chr(34)
            .Formulas(8) = "CarType=" & Chr(34) & Replace(tmpManifest.TruckTypeName & ", " & tmpManifest.TruckSubTypeName, vbCrLf, "") & Chr(34)
            .Formulas(9) = "TelNo=" & Chr(34) & Replace(tmpManifest.TelNo, vbCrLf, "") & Chr(34)
            .Formulas(10) = "DriverName=" & Chr(34) & Replace(tmpManifest.DriverName, vbCrLf, "") & Chr(34)
            .Formulas(11) = "AppointDate=" & Chr(34) & Replace(FormatYMD_to_DMY(selItem.AppointDate, "/", "/"), vbCrLf, "") & Chr(34)
            .Formulas(12) = "WorkerCount=" & Chr(34) & Replace(selItem.WorkerCount, vbCrLf, "") & Chr(34)
            .Formulas(13) = "TimeBack=" & Chr(34) & Replace(tmpManifest.OilUse, vbCrLf, "") & Chr(34)
            .Formulas(14) = "TimeIn=" & Chr(34) & Replace(tmpManifest.TimeArriveCust, vbCrLf, "") & Chr(34)
            .Formulas(14) = "TimeIn=" & Chr(34) & Replace(tmpManifest.StartTime, vbCrLf, "") & Chr(34)
            .Formulas(15) = "TimeOut=" & Chr(34) & Replace(tmpManifest.TimeOutCust, vbCrLf, "") & Chr(34)
            .Formulas(17) = "SaleStaff=" & Chr(34) & Replace("เจ้าหน้าที่การตลาด คุณ" & tmpManifest.UnderStaffName & "   เบอร์โทร. " & selItem.StTel, vbCrLf, "") & Chr(34)
            
            
            
        WasteCount = 1
         Do Until tmpWaste.EOF
            .Formulas(23 + WasteCount) = "WasteName" & WasteCount & "=" & Chr(34) & CStr(WasteCount) & ". " & Trim(tmpWaste!CustomerCode) & " " & Trim(tmpWaste!WasteName) & Chr(34)
            tmpWaste.MoveNext
            WasteCount = WasteCount + 1
            If WasteCount >= 10 Then
                 tmpWaste.MoveLast
                 tmpWaste.MoveNext
            End If
        Loop

'            Dim Sype As String
'            If tmpManifest.Tankcount <> "" Or tmpManifest.Tankcount > 0 Then
'                    Sype = Sype & "ถ.กว้าง" & " " & tmpManifest.Tankcount & "ใบ"
'            End If
'            If tmpManifest.BagType = "C" Or tmpManifest.BagType = "P" Then
'                    Sype = Sype & "ถุงบิกแบ็ค" & " " & tmpManifest.BagCount & "ใบ"
'            End If
'            If tmpManifest.TankcountS <> "" Or tmpManifest.TankcountS > 0 Then
'                    Sype = Sype & "ถ.จู๋" & " " & tmpManifest.TankcountS & "ใบ"
'            End If
            .Formulas(18) = "Remmark=" & Chr(34) & Replace(tmpManifest.CarRem, vbCrLf, "") & Chr(34)
            .Formulas(19) = "CommentTimeNo=" & Chr(34) & Replace(tmpManifest.CommentTimeNo, vbCrLf, "") & Chr(34)
            .Formulas(21) = "Rate=" & Chr(34) & Replace("เรทน้ำมัน " & tmpManifest.CarRate & " ลิตร", vbCrLf, "") & Chr(34)
            .Formulas(22) = "RemPhoto=" & Chr(34) & Replace(Photo & " " & PhoDesW & " " & PhoDesB & " " & PhoDesBF & " " & CustPhoto & " " & CustCon & " " & CloseWeight & " " & EmpTrain & " " & CarCheck & " " & NewCustJob, vbCrLf, "") & Chr(34)
        con = ""
        If tmpManifest.Tankcount = "" Then tmpManifest.Tankcount = "0"
        If tmpManifest.Tankcount > 0 Then
        con = "ถังปากกว้าง " & tmpManifest.Tankcount & " "
        End If
        If tmpManifest.TankTypeSCount = "" Then tmpManifest.TankTypeSCount = "0"
        If tmpManifest.TankTypeSCount > 0 Then
        con = con & "ถังปากจู๋ " & tmpManifest.TankTypeSCount & " "
        End If
        If tmpManifest.BagCount = "" Then tmpManifest.BagCount = "0"
        If tmpManifest.BagCount > 0 Then
        con = con & "ถุงเปิดปาก " & tmpManifest.BagCount & " "
        End If
        If tmpManifest.BagBack = "" Then tmpManifest.BagBack = "0"
        If tmpManifest.BagBack > 0 Then
        con = con & "ถุงดำ " & tmpManifest.BagBack & " "
        End If
        If tmpManifest.Container = "" Then tmpManifest.Container = "0"
        If tmpManifest.Container > 0 Then
        con = con & "เบ้าท์ใส่น้ำเสีย " & tmpManifest.Container & " "
        End If
        If tmpManifest.Lack = "" Then tmpManifest.Lack = "0"
        If tmpManifest.Lack > 0 Then
        con = con & "แลค " & tmpManifest.Lack & " "
        End If
        If tmpManifest.PaperBox = "" Then tmpManifest.PaperBox = "0"
        If tmpManifest.PaperBox > 0 Then
        con = con & "กล่องกระดาษ " & tmpManifest.PaperBox & " "
        End If
        If tmpManifest.BoutS = "" Then tmpManifest.BoutS = "0"
        If tmpManifest.BoutS > 0 Then
        con = con & "เบ้าท์ตัดปาก " & tmpManifest.BoutS & " "
        End If
        If tmpManifest.LackNew = "" Then tmpManifest.LackNew = "0"
        If tmpManifest.LackNew > 0 Then
        con = con & "แลครุ่นใหม่ " & tmpManifest.LackNew & " "
        End If
        If tmpManifest.LackGas = "" Then tmpManifest.LackGas = "0"
         If tmpManifest.LackGas > 0 Then
        con = con & "แลคใส่ถังแก๊ส " & tmpManifest.LackGas & " "
        End If
        If tmpManifest.Plate = "" Then tmpManifest.Plate = "0"
         If tmpManifest.Plate > 0 Then
        con = con & "พาเลท " & tmpManifest.Plate & " "
        End If
        If tmpManifest.Plastic = "" Then tmpManifest.Plastic = "0"
         If tmpManifest.Plastic > 0 Then
        con = con & "พลาสติกแล็ป " & tmpManifest.Plastic & " "
        End If
        If tmpManifest.BoutCage = "" Then tmpManifest.BoutCage = "0"
         If tmpManifest.BoutCage > 0 Then
        con = con & "กรงเบ้าท์ " & tmpManifest.BoutCage & " "
        End If
        If tmpManifest.CovidBin = "" Then tmpManifest.CovidBin = "0"
        If tmpManifest.CovidBin > 0 Then
        con = con & "ถังขยะโควิด " & tmpManifest.CovidBin & " "
        End If

            .Formulas(23) = "Con=" & Chr(34) & Replace(con, vbCrLf, "") & Chr(34)

            For x = 1 To tmpCol.Count
                Set tmpManifest = tmpCol(x)
                tmpStr = tmpStr & tmpManifest.DocNo & ", "
                If InStr(1, allCarRegisNo, tmpManifest.CarRegisID) = 0 Then
                    allCarRegisNo = allCarRegisNo & tmpManifest.CarRegisID & ", "
                End If
                Set tmpManifest = Nothing
            Next
            allCarRegisNo = Left(Trim(allCarRegisNo), Len(Trim(allCarRegisNo)) - 1)
            .Formulas(6) = "CarRegisID=" & Chr(34) & Replace(allCarRegisNo, vbCrLf, "") & Chr(34)
            
            If Trim(tmpStr) <> "" Then
                tmpStr = Left(Trim(tmpStr), Len(Trim(tmpStr)) - 1)
                .Formulas(16) = "MenifestList=" & Chr(34) & selItem.TimeTableNo & " : " & tmpStr & Chr(34)
            End If

            .Destination = crptToPrinter
            .PrintReport
            If Trim(.LastErrorString) <> "" Then MsgBox .LastErrorString
            
            .Formulas(0) = "ShowDay=" & Chr(34) & "" & Chr(34)
            .Formulas(1) = "ShowMonth=" & Chr(34) & "" & Chr(34)
            .Formulas(2) = "ShowYear=" & Chr(34) & "" & Chr(34)
            .Formulas(3) = "CompanyName=" & Chr(34) & "" & Chr(34)
            .Formulas(4) = "Vendor=" & Chr(34) & "" & Chr(34)
            .Formulas(5) = "ShowAddr=" & Chr(34) & "" & Chr(34)
            .Formulas(6) = "CarRegisID=" & Chr(34) & "" & Chr(34)
            .Formulas(7) = "ContactPerson=" & Chr(34) & "" & Chr(34)
            .Formulas(8) = "CarType=" & Chr(34) & "" & Chr(34)
            .Formulas(9) = "TelNo=" & Chr(34) & "" & Chr(34)
            .Formulas(10) = "DriverName=" & Chr(34) & "" & Chr(34)
            .Formulas(11) = "AppointDate=" & Chr(34) & "" & Chr(34)
            .Formulas(12) = "WorkerCount=" & Chr(34) & "" & Chr(34)
            .Formulas(13) = "TimeBack=" & Chr(34) & "" & Chr(34)
            .Formulas(14) = "TimeIn=" & Chr(34) & "" & Chr(34)
            .Formulas(15) = "TimeOut=" & Chr(34) & "" & Chr(34)
            .Formulas(16) = "MenifestList=" & Chr(34) & "" & Chr(34)
            .Formulas(17) = "SaleStaff=" & Chr(34) & "" & Chr(34)
            .Formulas(18) = "Remmark=" & Chr(34) & "" & Chr(34)
    '        .Reset
        End With
    End If
End Sub

Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If SaleStaff_ Is Nothing Then Set SaleStaff_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(SaleStaff_.StaffFName & " " & SaleStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set SaleStaff_ = Nothing
        Set SaleStaff_ = New BWGCompanyStaff
    Else
        If Trim(SaleStaff_.StaffFName & " " & SaleStaff_.StaffLName) <> Trim(txtSale.Text) Then
            Set SaleStaff_ = Nothing
            Set SaleStaff_ = New BWGCompanyStaff
            Load frmStaffSearch
            With frmStaffSearch
                .getCriteria (txtSale.Text)
                .SetCallerForm Me
                .Show 1
            End With
            If SaleStaff_.id = "" Then txtSale.Text = ""
        End If
    End If
End Sub

Public Sub SetDataStaff(selStaff As BWGCompanyStaff, Optional isAssigned As Boolean = True)
Dim tmpRec As New ADODB.Recordset
Dim x%
    Set SaleStaff_ = Nothing
    Set SaleStaff_ = selStaff
    If SaleStaff_ Is Nothing Then Exit Sub
    With SaleStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub

Public Sub UpdateTimeTableNo(TimeTableNo As String, CIndex As Long)
    lvResult.SelectedItem.SubItems(CIndex) = TimeTableNo
End Sub
