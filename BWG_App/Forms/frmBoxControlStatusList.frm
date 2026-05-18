VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBoxControlStatusList 
   BackColor       =   &H00FFFFFF&
   Caption         =   "สถานะใบ Menifest - F062"
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
   Begin Crystal.CrystalReport RPT3 
      Left            =   -90
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdTrans 
      Caption         =   "ใบกำกับการขนส่ง"
      Height          =   435
      Left            =   180
      TabIndex        =   28
      Top             =   5715
      Width           =   2295
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   3390
      Style           =   2  'Dropdown List
      TabIndex        =   26
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
      Height          =   9420
      Left            =   0
      ScaleHeight     =   9390
      ScaleWidth      =   2655
      TabIndex        =   12
      Top             =   0
      Width           =   2685
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   3360
         Width           =   2295
      End
      Begin VB.CommandButton cmdChangeTripNo 
         Caption         =   "เปลี่ยนเลขที่ใบคุม"
         Height          =   435
         Left            =   180
         TabIndex        =   9
         Top             =   6180
         Width           =   2295
      End
      Begin VB.TextBox txtJobNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   3960
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-"
         Height          =   255
         Left            =   780
         TabIndex        =   23
         Top             =   4380
         Width           =   555
      End
      Begin VB.CommandButton Command3 
         Caption         =   "+"
         Height          =   255
         Left            =   180
         TabIndex        =   22
         Top             =   4380
         Width           =   555
      End
      Begin VB.CommandButton cmdPrintBoxControl 
         Caption         =   "พิมพ์ใบคุม"
         Height          =   435
         Left            =   180
         TabIndex        =   8
         Top             =   5220
         Width           =   2295
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
         Top             =   1950
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   180
         TabIndex        =   10
         Top             =   6660
         Width           =   2295
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   180
         TabIndex        =   7
         Top             =   4740
         Width           =   2295
      End
      Begin VB.ComboBox cboJobStatus 
         Height          =   315
         ItemData        =   "frmBoxControlStatusList.frx":0000
         Left            =   180
         List            =   "frmBoxControlStatusList.frx":0002
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
         Top             =   2610
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Line Line2 
         X1              =   180
         X2              =   2490
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   27
         Top             =   3120
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลขที่ใบสั่งงาน,เลขที่ใบคุม :"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   24
         Top             =   3720
         Width           =   1890
      End
      Begin VB.Line Line1 
         X1              =   180
         X2              =   2490
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor, คนขับ, ทะเบียนรถ:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   8445
         Width           =   480
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปิดงานแล้ว"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   450
         TabIndex        =   19
         Top             =   8145
         Width           =   780
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พิมพ์ Manifest แล้ว"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   450
         TabIndex        =   18
         Top             =   7845
         Width           =   1335
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "งานใหม่"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   450
         TabIndex        =   17
         Top             =   7545
         Width           =   570
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   180
         Top             =   8415
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   180
         Top             =   8115
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   180
         Top             =   7815
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   180
         Top             =   7515
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   16
         Top             =   2370
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "วันที่ปฏิบัติงานตั้งแต่"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   15
         Top             =   1710
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   13
         Top             =   135
         Width           =   660
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2640
      TabIndex        =   11
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ใบสั่งงาน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ใบคุม"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "เที่ยวที่"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ลูกค้า"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "รถที่ใช้ขน"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "เจ้าของรถ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "คนขับ"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "วันที่ขน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "สถานะ"
         Object.Width           =   2117
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
            Picture         =   "frmBoxControlStatusList.frx":0004
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBoxControlStatusList.frx":059E
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
      TabIndex        =   25
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
      TabIndex        =   14
      Top             =   4950
      Width           =   1440
   End
End
Attribute VB_Name = "frmBoxControlStatusList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCondition_ As String
Dim colControlDoc_ As Collection
Dim PrintSelected As Boolean

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
        selTripNo = Trim(lvResult.SelectedItem.SubItems(1))
        .SetTripNo selTripNo
        .SetCallerForm Me
        .JobDataTimeTable "tbJobDataTimeTableBox"
        .Show 1
        If isExist(colControlDoc_, "'" & selTripNo & "'") Then
            colControlDoc_("'" & selTripNo & "'").TimeTableNo = lvResult.SelectedItem.SubItems(1)
        End If
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    SearchQue
    If lvResult.ListItems.Count = 0 Then MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
End Sub

Private Sub cmdViewJobBill_Click()
    lvResult_DblClick
End Sub

Private Sub cmdPrintBoxControl_Click()
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
        Set tmpCol = tmpS.JobDataTimeTableBoxSearch("TimeTableNo='" & colControlDoc_(x).TimeTableNo & "' And isMain='Y' AND ISNULL(isCancel, 'N') ='N' ")
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
        sqlStr = "SELECT * FROM vw_JobDataTimeTableBoxSearch WHERE isCancel<> 'Y' and  TimeTableNo='" & colControlDoc_(x).TimeTableNo & "'   ORDER BY companyName    "
        Set tmpRec = New ADODB.Recordset
        Set tmpCol = New Collection
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        With RPT3
            .LogonInfo(0) = "DSN = " & DSN_NAME & ";UID = " & strUid & ";PWD = " & strPwd
            .ReportFileName = App.Path & "\reports\TripMenifestDet2.rpt"
              If tmpRec(46) = "รถ Roll off Truck- พ่วง" Or tmpRec(0) = "รถ Lugger Box - พ่วง" Then
                Dim sqlrs1 As String
                sqlrs1 = " SELECT distinct CarRegisID,TruckTypeDesc,HeadorTrail FROM vw_MenifestDetInTimeNo WHERE isCanceled<> 'Y' and HeadorTrail <> 'T'  and  TimeTableNo='" & colControlDoc_(x).TimeTableNo & "'   ORDER BY HeadorTrail  "
                Set tmpRegisH = New ADODB.Recordset
               Set tmpRegisH = GetRS(sqlrs1, adOpenForwardOnly, adLockReadOnly)
                    If tmpRegisH.EOF Then
                        .Formulas(0) = "CarRegisH=" & Chr(34) & "" & Chr(34)
                        .Formulas(1) = "TypeHdesc=" & Chr(34) & "" & Chr(34)
                        .Formulas(2) = "CarRegisT=" & Chr(34) & "" & Chr(34)
                        .Formulas(3) = "TypeTdesc=" & Chr(34) & "" & Chr(34)
                    Else
                        .Formulas(0) = "CarRegisH=" & Chr(34) & tmpRegisH(0) & Chr(34)
                        .Formulas(1) = "TypeHdesc=" & Chr(34) & tmpRegisH(1) & Chr(34)
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
        Case 8
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
Dim JobID() As String
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    JobID = Split(lvResult.SelectedItem.Key, "_")
    Set tmpCol = tmpS.JobDataHeaderBoxSearch("JobID = " & JobID(0) & "'")
    If tmpCol Is Nothing Then Exit Sub
    If tmpCol.Count > 0 Then
        Load frmJobCarBoxAssign
        frmJobCarBoxAssign.SetDataJobOpen tmpCol(1)
        frmJobCarBoxAssign.Show 1
    End If
        Set tmpCol = Nothing
        Set tmpS = Nothing
End Sub

Private Sub lvResult_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim tmpItem As BWGJobDataTimeTableBox
    If colControlDoc_ Is Nothing Then Set colControlDoc_ = New Collection
    If item.Checked Then
        If isExist(colControlDoc_, "'" & item.SubItems(1) & "'") = False Then
            Set tmpItem = New BWGJobDataTimeTableBox
            tmpItem.TimeTableNo = item.SubItems(1)
            colControlDoc_.Add tmpItem, "'" & item.SubItems(1) & "'"
        End If
    Else
        If isExist(colControlDoc_, "'" & item.SubItems(1) & "'") Then
            colControlDoc_.Remove "'" & item.SubItems(1) & "'"
        End If
    End If
End Sub

Private Sub lvResult_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
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
                                Pont.x, Pont.Y, Me.hwnd, ByVal 0&)
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
Dim x As Long, Y As Long
Dim sqlStr As String, strCond As String, StatusDesc As String
Dim tmpArr
Dim curColor
Dim subStr As String
    Screen.MousePointer = 11
    Set colControlDoc_ = Nothing
    Set colControlDoc_ = New Collection
    sqlStr = "SELECT DISTINCT *" & vbCrLf
    sqlStr = sqlStr & "FROM vw_JobDataTimeTableBoxSearch "
    If Trim(txtJobNo.Text) <> "" Then
        strCond = " WHERE JobNo Like '%" & Trim(txtJobNo.Text) & "%' Or TimeTableNo Like '%" & Trim(txtJobNo.Text) & "%'"
    Else
        strCond = " WHERE (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "') "
        If Trim(txtKey.Text) <> "" And Replace(Trim(txtKey.Text), "*", "") <> "" Then
            strCond = strCond & " AND (CompanyName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')"
        End If
        If Trim(txtKey2.Text) <> "" And Replace(Trim(txtKey2.Text), "*", "") <> "" Then
            strCond = strCond & " AND (NameTH Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or DriverName Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%' Or CarRegisID Like '%" & Replace(Trim(txtKey2.Text), "*", "%") & "%')"
        End If
        Select Case cboJobStatus.ListIndex
            Case 1 'งานใหม่
                subStr = " ((ISNULL(isPrinted, 'N') <> 'Y') AND (ISNULL(isCancel, 'N') <> 'Y') AND (ISNULL(isCompleted, 'N') <> 'Y')) "
            Case 2 'พิมพ์ ใบคุม แล้ว
                subStr = " (isPrinted='Y' AND (ISNULL(isCancel, 'N') <> 'Y') AND (ISNULL(isCompleted, 'N') <> 'Y')) "
            Case 3 'ปิดงาน
                subStr = " (isCompleted = 'Y' AND (ISNULL(isCancel, 'N') <> 'Y')) "
            Case 4 'ยกเลิก
                subStr = " (isCancel = 'Y') "
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
        sqlStr = sqlStr & vbCrLf & "AND UnderSaleStaffID = '" & SaleStaff_.ID & "'" & vbCrLf
    End If
    
    sqlStr = sqlStr & " ORDER BY WorkDate, TimeTableNo, TripNo"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    
    
    If tmpRec Is Nothing Then
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ"
        Exit Sub
    End If
    lvResult.ListItems.Clear
    Do While Not tmpRec.EOF
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpRec!JobID & "_" & tmpRec!JobDataCarID & "_" & tmpRec!TripNo & "'", "" & tmpRec!JobNo)
        If Trim("" & tmpRec!isPrinted) <> "Y" And Trim("" & tmpRec!isCancel) <> "Y" And Trim("" & tmpRec!isCompleted) <> "Y" Then
            curColor = lbOpen.ForeColor
            StatusDesc = lbOpen.Caption
        ElseIf Trim("" & tmpRec!isPrinted) = "Y" And Trim("" & tmpRec!isCancel) <> "Y" And Trim("" & tmpRec!isCompleted) <> "Y" Then
            curColor = lbApproved.ForeColor
            StatusDesc = lbApproved.Caption
        ElseIf Trim("" & tmpRec!isCompleted) = "Y" And Trim("" & tmpRec!isCancel) <> "Y" Then
            curColor = lbClosed.ForeColor
            StatusDesc = lbClosed.Caption
        ElseIf Trim("" & tmpRec!isCancel) = "Y" Then
            curColor = lbCancel.ForeColor
            StatusDesc = lbCancel.Caption
        End If
        tmpList.ForeColor = curColor
        tmpList.SubItems(1) = Trim("" & tmpRec!TimeTableNo)
                tmpList.SubItems(2) = Trim("" & tmpRec!TripNo)
        tmpList.SubItems(3) = Trim("" & tmpRec!CustomerCode & ", " & tmpRec!CompanyName)
        If Trim("" & tmpRec!isTrailer) = "Y" Then
            tmpList.SubItems(4) = Trim("" & tmpRec!TruckTypeDesc) & " (พ่วง), " & Trim("" & tmpRec!SubTypeDesc)
        Else
            tmpList.SubItems(4) = Trim("" & tmpRec!TruckTypeDesc) & ", " & Trim("" & tmpRec!SubTypeDesc)
        End If
        tmpList.SubItems(5) = Trim("" & tmpRec!NameTH)
        tmpList.SubItems(6) = Trim("" & tmpRec!DriverName)
        tmpList.SubItems(7) = FormatYMD_to_DMY(Trim("" & tmpRec!WorkDate), "/", "/")
        tmpList.SubItems(8) = StatusDesc
        
        For Y = 1 To tmpList.ListSubItems.Count
            tmpList.ListSubItems(Y).ForeColor = curColor
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

Private Sub PrintDocControl(selItem As BWGJobDataTimeTableBox)
Dim x%, Y%, z%, iCount%, tmpCol As Collection
Dim tmpManifest As BWGMenifest, tmpStr As String, allCarRegisNo$
Dim UserGrp As Collection
Dim Group As New AWSUserGroup
Dim StrGrp As String
Dim i As Integer
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
        Set ExptoFile = New Scripting.FileSystemObject
        ExpReportPath = "\\192.168.1.2\ftproot\ExportManifest\" & Format(selItem.WorkDate, "yyyyMMdd")
        If ExptoFile.FolderExists(ExpReportPath) = False Then
            Set ExportFld = ExptoFile.CreateFolder(ExpReportPath)
        End If
    
        myCRReport.FormulaFields.GetItemByName("ShowDay").Text = Chr(34) & Right(selItem.WorkDate, 2) & Chr(34)
        myCRReport.FormulaFields.GetItemByName("ShowMonth").Text = Chr(34) & Mid(selItem.WorkDate, 6, 2) & Chr(34)
        myCRReport.FormulaFields.GetItemByName("ShowYear").Text = Chr(34) & Left(selItem.WorkDate, 4) & Chr(34)
        myCRReport.FormulaFields.GetItemByName("CompanyName").Text = Chr(34) & Replace(selItem.CustomerName, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("Vendor").Text = Chr(34) & Replace(selItem.VendorName, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("ShowAddr").Text = Chr(34) & Replace(selItem.SiteAddress, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("ContactPerson").Text = Chr(34) & Replace(selItem.ContactName, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("CarType").Text = Chr(34) & Replace(selItem.TruckTypeDesc & ", " & selItem.TruckSubTypeDesc, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("TelNo").Text = Chr(34) & Replace(selItem.TelNo, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("DriverName").Text = Chr(34) & Replace(selItem.DriverName, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("AppointDate").Text = Chr(34) & Replace(FormatYMD_to_DMY(selItem.AppointDate, "/", "/"), vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("WorkerCount").Text = Chr(34) & Replace(selItem.WorkerCount, vbCrLf, "") & Chr(34)
'        myCRReport.FormulaFields.GetItemByName("TimeBack").Text = Chr(34) & Replace(selItem.TimeArriveSite, vbCrLf, "") & Chr(34)
'        myCRReport.FormulaFields.GetItemByName("TimeIn").Text = Chr(34) & Replace(tmpManifest.TimeArriveCust, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("TimeIn").Text = Chr(34) & Replace(selItem.StartTime, vbCrLf, "") & Chr(34)
'        myCRReport.FormulaFields.GetItemByName("TimeOut").Text = Chr(34) & Replace(selItem.TimeOutCust, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("SaleStaff").Text = Chr(34) & Replace("เจ้าหน้าที่การตลาด คุณ" & selItem.SaleStaffName & "   เบอร์โทร. " & selItem.StaffTelNo, vbCrLf, "") & Chr(34)
        myCRReport.FormulaFields.GetItemByName("Remmark").Text = Chr(34) & Replace(selItem.CarRem, vbCrLf, "") & Chr(34)
        allCarRegisNo = selItem.CarRegisNo & " - " & selItem.TrailerTruckNo
        myCRReport.FormulaFields.GetItemByName("CarRegisID").Text = Chr(34) & Replace(allCarRegisNo, vbCrLf, "") & Chr(34)
        
        myCRReport.FormulaFields.GetItemByName("MenifestList").Text = Chr(34) & selItem.TimeTableNo & " : " & tmpStr & Chr(34)
    
        myCRReport.ExportOptions.FormatType = crEFTCrystalReport
        myCRReport.ExportOptions.DestinationType = crEDTDiskFile
        TripControlFile = ExpReportPath & "\" & selItem.TimeTableNo & "_" & Replace(selItem.CustomerName & "_" & allCarRegisNo, Chr(13), "") & ".rpt"
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
            .Formulas(3) = "CompanyName=" & Chr(34) & Replace(selItem.CustomerName, vbCrLf, "") & Chr(34)
            .Formulas(4) = "Vendor=" & Chr(34) & Replace(selItem.VendorName, vbCrLf, "") & Chr(34)
            .Formulas(5) = "ShowAddr=" & Chr(34) & Replace(selItem.SiteAddress, vbCrLf, "") & Chr(34)
            .Formulas(7) = "ContactPerson=" & Chr(34) & Replace(selItem.ContactName, vbCrLf, "") & Chr(34)
            .Formulas(8) = "CarType=" & Chr(34) & Replace(selItem.TruckTypeDesc & ", " & selItem.TruckSubTypeDesc, vbCrLf, "") & Chr(34)
            .Formulas(9) = "TelNo=" & Chr(34) & Replace(selItem.TelNo, vbCrLf, "") & Chr(34)
            .Formulas(10) = "DriverName=" & Chr(34) & Replace(selItem.DriverName, vbCrLf, "") & Chr(34)
            .Formulas(11) = "AppointDate=" & Chr(34) & Replace(FormatYMD_to_DMY(selItem.AppointDate, "/", "/"), vbCrLf, "") & Chr(34)
            .Formulas(12) = "WorkerCount=" & Chr(34) & Replace(selItem.WorkerCount, vbCrLf, "") & Chr(34)
            '.Formulas(13) = "TimeBack=" & Chr(34) & Replace(selItem.tim, vbCrLf, "") & Chr(34)
            '.Formulas(14) = "TimeIn=" & Chr(34) & Replace(selItem.TimeArriveCust, vbCrLf, "") & Chr(34)
            .Formulas(14) = "TimeIn=" & Chr(34) & Replace(selItem.StartTime, vbCrLf, "") & Chr(34)
'            .Formulas(15) = "TimeOut=" & Chr(34) & Replace(selItem.Timeout, vbCrLf, "") & Chr(34)
            .Formulas(17) = "SaleStaff=" & Chr(34) & Replace("เจ้าหน้าที่การตลาด คุณ" & selItem.SaleStaffName & "   เบอร์โทร. " & selItem.StaffTelNo, vbCrLf, "") & Chr(34)
            .Formulas(18) = "Remmark=" & Chr(34) & Replace(selItem.CarRem, vbCrLf, "") & Chr(34)
            allCarRegisNo = selItem.CarRegisNo & " - " & selItem.TrailerTruckNo
            .Formulas(6) = "CarRegisID=" & Chr(34) & Replace(allCarRegisNo, vbCrLf, "") & Chr(34)
            .Formulas(16) = "MenifestList=" & Chr(34) & selItem.TimeTableNo & " : " & tmpStr & Chr(34)
    
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

Public Sub UpdateTimeTableNo(TimeTableNo As String, CIndex As Long)
    lvResult.SelectedItem.SubItems(CIndex) = TimeTableNo
End Sub
