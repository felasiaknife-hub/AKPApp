VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenifestClosedList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ใบ Menifest ที่ปิดงานแล้ว - F059"
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
   Begin VB.Frame fraMessage 
      Height          =   675
      Left            =   4740
      TabIndex        =   31
      Top             =   1980
      Visible         =   0   'False
      Width           =   3135
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "กำลังประมวลผล ห้ามปิดโปรแกรม"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   2655
      End
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
      Height          =   9420
      Left            =   0
      ScaleHeight     =   9390
      ScaleWidth      =   2355
      TabIndex        =   15
      Top             =   0
      Width           =   2385
      Begin VB.ComboBox cboCustType 
         Height          =   315
         Left            =   180
         TabIndex        =   33
         Top             =   2520
         Width           =   1815
      End
      Begin VB.ComboBox cboCarType 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2160
         Width           =   2025
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   4440
         Width           =   1995
      End
      Begin VB.TextBox txtTimeTableNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   8
         Top             =   5100
         Width           =   1995
      End
      Begin VB.CommandButton aaaa 
         Caption         =   "Command1"
         Height          =   315
         Left            =   330
         TabIndex        =   27
         Top             =   9060
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.ComboBox cboStatus3 
         Height          =   315
         ItemData        =   "frMenifestClosedList.frx":0000
         Left            =   180
         List            =   "frMenifestClosedList.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1800
         Width           =   2025
      End
      Begin VB.ComboBox cboStatus2 
         Height          =   315
         ItemData        =   "frMenifestClosedList.frx":0004
         Left            =   180
         List            =   "frMenifestClosedList.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   2025
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-"
         Height          =   255
         Left            =   660
         TabIndex        =   10
         Top             =   5490
         Width           =   555
      End
      Begin VB.CommandButton Command3 
         Caption         =   "+"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   5490
         Width           =   555
      End
      Begin VB.TextBox txtKey2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Text            =   "*"
         Top             =   720
         Width           =   1995
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   5
         Top             =   3240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.CommandButton cmdViewJobBill 
         Caption         =   "ดูข้อมูลใบ Manifest"
         Height          =   435
         Left            =   60
         TabIndex        =   12
         Top             =   6240
         Width           =   2205
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   60
         TabIndex        =   13
         Top             =   6690
         Width           =   2205
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   60
         TabIndex        =   11
         Top             =   5790
         Width           =   2205
      End
      Begin VB.ComboBox cboJobStatus 
         Height          =   315
         ItemData        =   "frMenifestClosedList.frx":0008
         Left            =   180
         List            =   "frMenifestClosedList.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   2025
      End
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   0
         Text            =   "*"
         Top             =   240
         Width           =   1995
      End
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   6
         Top             =   3840
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   29
         Top             =   4200
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบคุม"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   28
         Top             =   4860
         Width           =   765
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   90
         X2              =   2220
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   5
         Left            =   120
         Top             =   7500
         Width           =   195
      End
      Begin VB.Label lbHalfCheck 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตรวจแล้วเอกสารไม่ครบ"
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   390
         TabIndex        =   26
         Top             =   7530
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00000040&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   3
         Left            =   120
         Top             =   8430
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ทำจ่าย Vendor แล้ว"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   390
         TabIndex        =   25
         Top             =   8460
         Width           =   1380
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   90
         X2              =   2220
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor, คนขับ, ทะเบียนรถ:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   24
         Top             =   540
         Width           =   1965
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   390
         TabIndex        =   23
         Top             =   8790
         Width           =   480
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วางบิลแล้ว"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   390
         TabIndex        =   22
         Top             =   8130
         Width           =   735
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตรวจสอบเอกสารครบแล้ว"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   390
         TabIndex        =   21
         Top             =   7830
         Width           =   1785
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยังไม่ได้ตรวจสอบ"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   390
         TabIndex        =   20
         Top             =   7230
         Width           =   1230
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   120
         Top             =   8760
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   120
         Top             =   8100
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   120
         Top             =   7800
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   120
         Top             =   7200
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   19
         Top             =   3600
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ปฏิบัติงานตั้งแต่"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   3000
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลข Manifest, ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   16
         Top             =   60
         Width           =   1650
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2370
      TabIndex        =   14
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
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
      NumItems        =   18
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
         Object.Width           =   2469
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
         Text            =   "สถานะ"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "วางบิล"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "จ่าย Vendor"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ใบงาน"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ใบกำกับ"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "ใบน้ำหนัก"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "ใบน้ำหนักลูกค้า"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "ใบคุม"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Invoice No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "ผู้ปิดงาน"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "เลขทะเบียนโรงงาน"
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
      Left            =   6000
      Top             =   3540
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
            Picture         =   "frMenifestClosedList.frx":000C
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frMenifestClosedList.frx":05A6
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
      TabIndex        =   17
      Top             =   4950
      Width           =   1440
   End
End
Attribute VB_Name = "frmMenifestClosedList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' In General Declarations
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_HITTEST As Long = (LVM_FIRST + 18)
Private Const LVM_SUBITEMHITTEST As Long = (LVM_FIRST + 57)
Private Const LVHT_ONITEMICON As Long = &H2
Private Const LVHT_ONITEMLABEL As Long = &H4
Private Const LVHT_ONITEMSTATEICON As Long = &H8
Private Const LVHT_ONITEM As Long = (LVHT_ONITEMICON Or _
                                    LVHT_ONITEMLABEL Or _
                                    LVHT_ONITEMSTATEICON)

'Private Type POINTAPI
'  x As Long
'  Y As Long
'End Type

Dim lX As Single, lY As Single

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim curCondition_ As String
Dim ColumnIndex As Integer

Dim SaleStaff_ As BWGCompanyStaff

Private Sub aaaa_Click()
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim iCount As Integer
Dim runNo As Long
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbJobDataTimeTable ORDER BY WorkDate"
    
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    DBConnExc "Delete tbRunTimeTableNo"
    DBConnExc "Update view_8 Set WorkDate=Expr1"
    runNo = 1
    GetDBConnection
    Do Until tmpRec.EOF
        iCount = 0
        DBConn.Execute "Update tbRunTimeTableNo Set AddYear='" & Left(tmpRec("WorkDate"), 4) & "' WHERE JobDataCarID='" & tmpRec("JobDataCarID") & "' And TripNo=" & tmpRec("TripNo"), iCount
        If iCount = 0 Then
            DBConnExc "Insert into tbRunTimeTableNo (JobDataCarID,TripNo,AddYear,RunNumber) Values ('" & tmpRec("JobDataCarID") & "'," & tmpRec("TripNo") & ",'" & Left(tmpRec("WorkDate"), 4) & "'," & runNo & ")"
            DBConnExc "Update tbJobDataTimeTable Set TimeTableNo='" & Left(tmpRec("WorkDate"), 4) & Format(runNo, "00000") & "' WHERE JobDataCarID='" & tmpRec("JobDataCarID") & "' And TripNo=" & tmpRec("TripNo")
            runNo = runNo + 1
        End If
        tmpRec.MoveNext
    Loop
MsgBox "Done"
CloseDBConnection
Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    frmMain.Stb1.Panels(1).Text = ""
    frmMain.Stb1.Panels(2).Text = ""
    SearchQue
    If lvResult.ListItems.Count = 0 Then MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
End Sub

Private Sub Command1_Click()
Dim x%, iCount%
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีรายการใบ Manifest ที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    For x = 1 To lvResult.ListItems.Count
        If lvResult.ListItems(x).Checked Then
            iCount = iCount + 1
        End If
    Next
    If iCount = 0 Then
        MsgBox "กรุณาเลือกรายการใบ Manifest ที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    
    MsgBox "เสร็จสิ้นการพิมพ์ใบ Manifest แล้วจำนวน " & iCount & " รายการ"
End Sub

Private Sub Command2_Click()
Dim x%, iCount%
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่มีรายการใบ Manifest ที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    For x = 1 To lvResult.ListItems.Count
        If lvResult.ListItems(x).Checked Then
            iCount = iCount + 1
        End If
    Next
    If iCount = 0 Then
        MsgBox "กรุณาเลือกรายการใบ Manifest ที่ต้องการพิมพ์", vbExclamation
        Exit Sub
    End If
    
    MsgBox "เสร็จสิ้นการพิมพ์ใบคุมแล้วจำนวน " & iCount & " รายการ"

End Sub

Private Sub cmdViewJobBill_Click()
    lvResult_DblClick
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
                If Me.ActiveControl.Name <> txtSale.Name Then cmdSearch_Click
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
Dim backDate As String
Dim tmpY As Integer, tmpM As Integer, tmpD As Integer
Dim tmpArr
Dim sqlStr As String
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
    ' Hook Form
'    Call WheelHook(Me.hWnd)
  
  ' Hook Controls to be ignored
    Dim ctl As Control
    For Each ctl In Controls
        On Error Resume Next
        Select Case True
        Case TypeOf ctl Is MSFlexGrid, TypeOf ctl Is VSFlexGrid
            Call WheelHook(ctl.hwnd)
        Case TypeOf ctl Is PictureBox
            Call WheelHook(ctl.hwnd)
        Case TypeOf ctl Is ListBox, TypeOf ctl Is ListView, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
            Call WheelHook(ctl.hwnd)
        Case Else
            '
        End Select
    Next ctl
    
    sqlStr = "select backdate=convert(varchar(10),getdate()-7,111)"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then backDate = "0"
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
        .AddItem "== ทุกสถานะเอกสาร =="
        .AddItem "ยังไม่ได้ตรวจสอบ"
        .AddItem "ตรวจแล้วเอกสารไม่ครบ"
        .AddItem "ตรวจแล้วเอกสารครบ"
        .ListIndex = 0
    End With
    With cboStatus2
        .AddItem "== ทุกสถานะวางบิล =="
        .AddItem "ยังไม่ได้วางบิล"
        .AddItem "วางบิลแล้ว"
        .AddItem "ไม่ได้วางบิล(วางบิลเลขอื่น)"
        .AddItem "ไม่ได้วางบิล(ไม่ได้ของ)"
        .ListIndex = 0
    End With
    With cboStatus3
        .AddItem "== ทุกสถานะจ่าย Vendor =="
        .AddItem "ยังไม่ได้จ่าย"
        .AddItem "จ่าย Vendor แล้ว"
        .ListIndex = 0
    End With
    PopulateCarType cboCarType
    With cboCustType
        .AddItem "== ประเภทลูกค้า =="
        .AddItem "(C) ,(s)"
        .AddItem "(F) RDF Waste"
        .AddItem "(L) (p)"
        .AddItem "(T) (k)"
        .AddItem "(G) (e)"
        .AddItem "(R)"
        .ListIndex = 0
    End With
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
    fraMessage.Top = (Me.Height / 2) - fraMessage.Height
    fraMessage.Left = (Me.Width / 2) - (fraMessage.Width / 2)
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
        Case 7
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
    On Error GoTo ErrD
    Dim myMenifest_ As BWGMenifest
    Dim tmpJobDataTimeTable As BWGJobDataTimeTable
    Dim tmpS As New BWGSearchManager
    Dim tmpCol As Collection
        
    Dim HTI As LVHITTESTINFO
    With HTI
        .Pt.x = (lX \ Screen.TwipsPerPixelX)
        .Pt.y = (lY \ Screen.TwipsPerPixelY)
        .Flags = LVHT_ONITEM
    End With
    
    Call SendMessage(lvResult.hwnd, LVM_SUBITEMHITTEST, 0, HTI)
    
    Dim lst As ListItem
    If HTI.iItem = -1 Then HTI.iItem = 0
    If (HTI.iItem > -1) Then
        Set lst = lvResult.ListItems(HTI.iItem + 1)
        If HTI.iSubItem = 0 Then
            ColumnIndex = HTI.iSubItem + 1
        Else
            ColumnIndex = HTI.iSubItem
        End If
'        MsgBox "Clicked item " & HTI.iItem + 1 & " and SubItem " & HTI.iSubItem
    End If
        
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    If ColumnIndex = 15 And lvResult.SelectedItem.SubItems(ColumnIndex) <> "" Then
        Set tmpCol = tmpS.BillingHeaderSearch("BillingHeaderID IN (SELECT InvID FROM tbMenifestHeader WHERE MenifestID = " & lvResult.SelectedItem.Key & ")")
        If tmpCol.Count > 0 Then
            Load frmInvoice
            frmInvoice.SetDataBillHeader tmpCol(1)
            frmInvoice.Show 1
        End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
    Else
        Set tmpCol = tmpS.MenifestSearch("MenifestID=" & lvResult.SelectedItem.Key)
        If tmpCol.Count > 0 Then
            If AutoCalcMF = False Then
                Load frmMenifest
                frmMenifest.SetDataMenifest tmpCol(1)
                frmMenifest.Show 1
            Else
'                Load frmMenifest
'                frmMenifest.SetDataMenifest tmpCol(1)
'                frmMenifest.CalcTrip
                Set myMenifest_ = tmpCol(1)
                Set tmpCol = tmpS.JobDataTimeTableQuickSearch("TimeTableNo='" & myMenifest_.TimeTableNo & "' And isMain='Y' And isCancel ='N' AND ISNULL(isCanceled, 'N') <> 'Y' AND TripNo = " & myMenifest_.TripNo)
                If tmpCol.Count > 0 Then
                    Set tmpJobDataTimeTable = tmpCol(1)
                    Load frmTripControl
                    With frmTripControl
                        .SetTripControl tmpJobDataTimeTable, myMenifest_.CustomerName, myMenifest_.TruckTypeName, myMenifest_.CarRegisID, myMenifest_.CustSiteAumphurID, myMenifest_.TSDFAumphurID, myMenifest_.TruckTypeID
'                        If AutoCalcMF = False Then
'                            .Show 1
'                        Else
                            frmTripControl.SaveData
                            frmMain.Stb1.Panels(1).Text = "จัดเก็บข้อมูลใบคุมเลขที่ " & myMenifest_.TimeTableNo & " ใบเมนิเฟสเลขที่ " & myMenifest_.DocNo
'                        End If
                    End With
                End If
            End If
        End If
        Set tmpJobDataTimeTable = Nothing
        Set myMenifest_ = Nothing
        Set tmpCol = Nothing
        Set tmpS = Nothing
    End If
    Exit Sub
ErrD:
        'GetMsg "", vbCritical
        If Not myMenifest_ Is Nothing Then
            MsgBox err.Number & " เกิดข้อผิดพลาด, ไม่สามารถบันทึกใบคุมเลขที่ " & myMenifest_.TimeTableNo & " ได้", vbCritical, "Critical Error"
        Else
            MsgBox err.Description, vbCritical, "Critical Error"
        End If
        Exit Sub
End Sub

Private Sub lvResult_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim Pont As POINTAPI
    Dim ret As BTN_STYLE

    If Button = vbRightButton Then
        hMenu = CreatePopupMenu()
        AppendMenu hMenu, MF_STRING, 1, "Auto Calculate Manifest"
        AppendMenu hMenu, MF_SEPARATOR, -1, "SEP1"
        AppendMenu hMenu, MF_STRING, 2, "แสดงเฉพาะรายการที่มียอดค่าบริการ = 0"
        AppendMenu hMenu, MF_STRING, 3, "แสดงเฉพาะรายการที่มีน้ำหนักรวม = 0"
        AppendMenu hMenu, MF_STRING, 4, "แสดงเฉพาะรายการที่ค่าขนส่งแฝง = 0"
        AppendMenu hMenu, MF_STRING, 5, "แสดงเฉพาะรายการที่ค่าขนส่งจ่ายเวนเดอร์ = 0"
'        If LCase(CurrentUser.LoginName) = "admin" Then AppendMenu hMenu, MF_STRING, 4, "อนุมัติจ่าย"
        GetCursorPos Pont
        ret = TrackPopupMenuEx(hMenu, TPM_LEFTALIGN Or TPM_RETURNCMD, _
                                Pont.x, Pont.y, Me.hwnd, ByVal 0&)
        DestroyMenu hMenu

        If ret = 1 Then
            MNU_AutoCalcMF
        ElseIf ret = 2 Then
            SearchQue False, False, False, True
        ElseIf ret = 3 Then
            SearchQue True, False, False, False
        ElseIf ret = 4 Then
            SearchQue False, True, False, False
        ElseIf ret = 5 Then
            SearchQue False, False, True, False
'        ElseIf ret = 4 Then
'            MNU_SetApprove_Click
        End If
    End If
End Sub

Private Sub lvResult_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   lX = x
   lY = y
End Sub

Private Sub MNU_AutoCalcMF()
'On Error GoTo ErrD
Dim i As Long, SelectedItem As ListItem, StartTime As Date
    If lvResult.ListItems.Count = 0 Then fraMessage.Visible = False: Exit Sub
    If MsgBox("Are you sure to auto calculate manifest document?", vbYesNo) = vbNo Then Exit Sub
    AutoCalcMF = True
    fraMessage.Visible = True
    StartTime = Format(Now, "dd/MM/yyyy HH:MM:SS")
    For i = 1 To lvResult.ListItems.Count
        DoEvents
        Set SelectedItem = lvResult.ListItems.item(i)
        lvResult.SelectedItem = SelectedItem
        lvResult_DblClick
        DoEvents
        frmMain.Stb1.Panels(2).Text = Format(i / lvResult.ListItems.Count * 100, "0.00") & "%"
    Next
    AutoCalcMF = False
    fraMessage.Visible = False
    frmMain.Stb1.Panels(1).Text = "Auto calculate completed. (" & i - 1 & ") records. (Start " & StartTime & " - End " & Format(Now, "dd/mm/yyyy HH:MM:SS") & ")"
    MsgBox "Auto calculate completed. (" & i - 1 & ") records."
    Exit Sub
ErrD:
    frmMain.Stb1.Panels(1).Text = err.Description
    frmMain.Stb1.Panels(2).Text = ""
    Resume Next
End Sub

Private Sub Timer1_Timer()
    'SearchQue
End Sub

Private Sub SearchQue(Optional ByVal ShowActualQty0 As Boolean = False, Optional ByVal ShowTripTransFee_Inc0 As Boolean = False, Optional ByVal ShowVDTransFee0 As Boolean = False, Optional ByVal ShowFee0 As Boolean = False)
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim tmpList As ListItem
Dim x As Long, y As Long
Dim sqlStr As String, strCond As String, StatusDesc As String
Dim tmpArr
Dim curColor
Dim subStr As String
    Screen.MousePointer = 11
    sqlStr = "Select MenifestID,DocNo,MenifestNo,CustomerName,SubTypeDesc,TruckTypeDesc,TransportName1,TransportName2,WorkDate,isClosed,isCanceled,isPrinted,DriverName,JobNo,DocJobRecord,DocControl,DocSiteWeight,DocCustWeight,DocCustControl,BillPayHeaderID,BillPayNo,BillingHeaderID,BillingNo,isDocCompleted, CASE WHEN ISNULL(isCancelHead, 'Y') = 'Y' THEN 'N' ELSE CASE WHEN ISNULL(BillingNo, '') = '' THEN ISNULL(isInv, 'N') ELSE 'Y' END END AS isInv, isCancelHead AS InvIsCancel " & vbCrLf
    sqlStr = sqlStr & ", TripTransFee_Inc, VDTransFee, CloseStaff = (SELECT TOP(1) StaffFName + ' ' + StaffLName FROM tbCompanyStaff WHERE CloseStaffID = StaffID),TimetableNo,[14Digit],RegisterNo" & vbCrLf
    sqlStr = sqlStr & "From vw_MenifestQue " & vbCrLf
    If Trim(txtTimeTableNo.Text) = "" Then
        strCond = " WHERE (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "') And (isClosed='Y') And (isCanceled<>'Y' Or isCanceled is null) "
        'strCond = " WHERE (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "') and MenifestNo<>'' "
        If Trim(txtKey.Text) <> "" And Replace(Trim(txtKey.Text), "*", "") <> "" Then
            strCond = strCond & " And (MenifestNo Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or DocNo Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or CustomerName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')"
        End If
        If Trim(txtKey2.Text) <> "" And Replace(Trim(txtKey2.Text), "*", "") <> "" Then
            strCond = strCond & " And (TransportName1 Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or TransportName2 Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or DriverName Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or CarRegisID Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%')"
        End If
        If cboCustType.ListIndex = 1 Then
            strCond = strCond & " AND CustTypeFlag in ('C','S') "
        ElseIf cboCustType.ListIndex = 2 Then
            strCond = strCond & " AND CustTypeFlag = 'F' "
        ElseIf cboCustType.ListIndex = 3 Then
            strCond = strCond & " AND CustTypeFlag in ('L','P') "
        ElseIf cboCustType.ListIndex = 4 Then
            strCond = strCond & " AND CustTypeFlag in ('T','K') "
        ElseIf cboCustType.ListIndex = 5 Then
            strCond = strCond & " AND CustTypeFlag in ('G','E') "
        ElseIf cboCustType.ListIndex = 6 Then
            strCond = strCond & " AND CustTypeFlag in ('R') "
        End If
        
        Select Case cboJobStatus.ListIndex
            Case 1 'ยังไม่ได้ตรวจสอบ
                subStr = " (DocJobRecord+DocControl+DocSiteWeight+DocCustWeight+DocCustControl='NNNNN') "
            Case 2 'ตรวจสอบเอกสารแล้ว
                subStr = " (CHARINDEX('Y',DocJobRecord+DocControl+DocSiteWeight+DocCustWeight+DocCustControl)>0 And CHARINDEX('N',DocJobRecord+DocControl+DocSiteWeight+DocCustWeight+DocCustControl)>0) "
            Case 3
                'subStr = " (DocJobRecord+DocControl+DocSiteWeight+DocCustWeight+DocCustControl='YYYYY') "
                subStr = " (isDocCompleted='Y') "
        End Select
        If cboJobStatus.ListIndex > 0 Then
            strCond = strCond & " And " & subStr
        End If
        Select Case cboStatus2.ListIndex
            Case 1 'ยังไม่ได้วางบิล
'                subStr = " (BillingHeaderID is null) "
                subStr = "(CASE WHEN ISNULL(isCancelHead, 'Y') = 'Y' THEN 'N' ELSE CASE WHEN ISNULL(BillingNo, '') = '' THEN ISNULL(isInv, 'N') ELSE 'Y' END END <> 'Y')  And ActualQty <> 0  "
            Case 2 'วางบิลแล้ว
'                subStr = " (BillingHeaderID is not null) "
                subStr = " (isInv = 'Y') "
            Case 3 'ยังไม่ได้วางบิลแต่มีการวางบิลจากเลขอื่น
                subStr = "(CASE WHEN ISNULL(isCancelHead, 'Y') = 'Y' THEN 'N' ELSE CASE WHEN ISNULL(BillingNo, '') = '' THEN ISNULL(isInv, 'N') ELSE 'Y' END END <> 'Y') "
                subStr = subStr & "  and InCollectFee = '0.00' AND TimetableNo in (Select TimetableNo from vw_BillingDetailSearch Where isCancel <> 'Y')  "
'                subStr = subStr & " AND MenifestID Not in (Select SellID from vw_BillingDetailSearch)  "
            Case 4
                subStr = "(CASE WHEN ISNULL(isCancelHead, 'Y') = 'Y' THEN 'N' ELSE CASE WHEN ISNULL(BillingNo, '') = '' THEN ISNULL(isInv, 'N') ELSE 'Y' END END <> 'Y')  And ActualQty  = 0 "
'                subStr = subStr & "  AND TimetableNo in (Select TimetableNo from vw_MenifestQue Group by TimetableNo,isCanceled having Sum(InCollectFee) = '0.00' and isCanceled <> 'Y')  "
                
        End Select
        
        If cboStatus2.ListIndex > 0 Then
            strCond = strCond & " And " & subStr
        End If
        Select Case cboStatus3.ListIndex
            Case 1 'ยังไม่ได้จ่าย
                subStr = " (BillPayHeaderID is null) "
            Case 2 'จ่ายแล้ว
                subStr = " (BillPayHeaderID is not null) "
        End Select
        If cboStatus3.ListIndex > 0 Then
            strCond = strCond & " And " & subStr
        End If
        
        If cboCarType.ListIndex > 0 Then strCond = strCond & " AND TruckTypeID='" & coltruckType_(cboCarType.ListIndex).ID & "'" & vbCrLf
        
        If ShowActualQty0 = True Then
            strCond = strCond & " AND (ActualQty = 0)"
        End If
        
        If ShowTripTransFee_Inc0 = True Then
            strCond = strCond & " AND TripTransFee_Inc = 0"
        End If
        
        If ShowVDTransFee0 = True Then
            strCond = strCond & " AND VDTransFee = 0"
        End If
        
        If ShowFee0 = True Then
            strCond = strCond & " AND ISNULL(InCollectFee, 0) = 0"
        End If
        
    Else
        strCond = " WHERE (TimeTableNo='" & txtTimeTableNo.Text & "') "
    End If
    
    curCondition_ = strCond
    
    sqlStr = sqlStr & strCond
    sqlStr = SelectWithCompanyID(sqlStr, False)
    
    If Trim(txtSale.Text) = "" Or Trim(txtSale.Text) = "*" Then Set SaleStaff_ = Nothing
    If Not SaleStaff_ Is Nothing Then
        sqlStr = sqlStr & vbCrLf & "AND UnderSaleStaffID = '" & SaleStaff_.ID & "'"
    End If
    
    sqlStr = sqlStr & vbCrLf & "ORDER BY WorkDate, DocNo, MenifestNo"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ  จำนวน 0 เที่ยว"
        Exit Sub
    End If
    lvResult.ListItems.Clear
    If tmpRec.EOF = False Then
        tmpArr = tmpRec.GetRows
        For x = 0 To UBound(tmpArr, 2)
            Set tmpList = lvResult.ListItems.Add(, "'" & tmpArr(0, x) & "'", "" & tmpArr(2, x) & ", " & Trim("" & tmpArr(1, x)))
            If Trim("" & tmpArr(14, x)) & Trim("" & tmpArr(15, x)) & Trim("" & tmpArr(16, x)) & Trim("" & tmpArr(17, x)) & Trim("" & tmpArr(18, x)) = "NNNNN" And Trim("" & tmpArr(23, x)) <> "Y" Then
                curColor = lbOpen.ForeColor
                StatusDesc = lbOpen.Caption
            ElseIf InStr(1, Trim("" & tmpArr(14, x)) & Trim("" & tmpArr(15, x)) & Trim("" & tmpArr(16, x)) & Trim("" & tmpArr(17, x)) & Trim("" & tmpArr(18, x)), "N") > 0 And InStr(1, Trim("" & tmpArr(14, x)) & Trim("" & tmpArr(15, x)) & Trim("" & tmpArr(16, x)) & Trim("" & tmpArr(17, x)) & Trim("" & tmpArr(18, x)), "Y") > 0 And Trim("" & tmpArr(23, x)) <> "Y" Then
                curColor = lbHalfCheck.ForeColor
                StatusDesc = lbHalfCheck.Caption
            'ElseIf Trim("" & tmpArr(14, x)) & Trim("" & tmpArr(15, x)) & Trim("" & tmpArr(16, x)) & Trim("" & tmpArr(17, x)) & Trim("" & tmpArr(18, x)) = "YYYYY" Then
            ElseIf Trim("" & tmpArr(23, x)) = "Y" Then
                curColor = lbApproved.ForeColor
                StatusDesc = lbApproved.Caption
            ElseIf Trim("" & tmpArr(21, x)) <> "" Then
                curColor = lbClosed.ForeColor
                StatusDesc = lbClosed.Caption
            ElseIf Trim("" & tmpArr(19, x)) <> "" Then
                curColor = Label1.ForeColor
                StatusDesc = Label1.Caption
            ElseIf Trim("" & tmpArr(10, x)) = "Y" Then
                curColor = lbCancel.ForeColor
                StatusDesc = lbCancel.Caption
            End If
            tmpList.ForeColor = curColor
            tmpList.SubItems(1) = Trim("" & tmpArr(3, x))
            tmpList.SubItems(2) = Trim("" & tmpArr(13, x))
            tmpList.SubItems(3) = Trim("" & tmpArr(5, x)) & ", " & Trim("" & tmpArr(4, x))
            tmpList.SubItems(4) = Trim("" & tmpArr(7, x))
            tmpList.SubItems(5) = Trim("" & tmpArr(12, x))
            tmpList.SubItems(6) = FormatYMD_to_DMY(Trim("" & tmpArr(8, x)), "/", "/")
            tmpList.SubItems(7) = StatusDesc
            If Trim("" & tmpArr(24, x)) = "Y" And Trim("" & tmpArr(25, x)) <> "Y" Then
                tmpList.SubItems(8) = "Y"
            ElseIf tmpArr(22, x) <> "" Then
                tmpList.SubItems(8) = "Y"
            Else
                tmpList.SubItems(8) = "N"
            End If
            If Trim("" & tmpArr(19, x)) <> "" Then
                tmpList.SubItems(9) = "Y"
            Else
                tmpList.SubItems(9) = "N"
            End If
            tmpList.SubItems(10) = Trim("" & tmpArr(14, x))
            tmpList.SubItems(11) = Trim("" & tmpArr(15, x))
            tmpList.SubItems(12) = Trim("" & tmpArr(16, x))
            tmpList.SubItems(13) = Trim("" & tmpArr(17, x))
            tmpList.SubItems(14) = Trim("" & tmpArr(18, x))
            tmpList.SubItems(15) = Trim("" & tmpArr(22, x))
            tmpList.SubItems(16) = Trim("" & tmpArr(28, x))
            tmpList.SubItems(17) = Trim("" & tmpArr(30, x))
            For y = 1 To tmpList.ListSubItems.Count
                tmpList.ListSubItems(y).ForeColor = curColor
            Next
            Set tmpList = Nothing
            'DoEvents
        Next
        Erase tmpArr
    End If
    lvResult.Sorted = False
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

Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
    Dim ctl As Control
    Dim bHandled As Boolean
    Dim bOver As Boolean
  
    For Each ctl In Controls
        ' Is the mouse over the control
        On Error Resume Next
        bOver = (ctl.Visible And IsOver(ctl.hwnd, Xpos, Ypos))
        On Error GoTo 0
    
        If bOver Then
            ' If so, respond accordingly
            bHandled = True
            Select Case True
            Case TypeOf ctl Is MSFlexGrid, TypeOf ctl Is VSFlexGrid
                FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos
        
            Case TypeOf ctl Is PictureBox
                'PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
            
            Case TypeOf ctl Is ListBox, TypeOf ctl Is ListView, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
                ' These controls already handle the mousewheel themselves, so allow them to:
                If ctl.Enabled Then ctl.SetFocus
    
            Case Else
            bHandled = False
        
            End Select
            If bHandled Then Exit Sub
        End If
        bOver = False
    Next ctl
    
    'Scroll was not handled by any controls, so treat as a general message send to the form
    'Me.Caption = "Form Scroll " & IIf(Rotation < 0, "Down", "Up")
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
            If SaleStaff_.ID = "" Then txtSale.Text = ""
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
