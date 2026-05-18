VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCarWorkingQue 
   BackColor       =   &H00FFFFFF&
   Caption         =   "สถานะรถ - F007"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   ForeColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      TabIndex        =   6
      Top             =   0
      Width           =   2385
      Begin VB.TextBox txtKey2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   180
         TabIndex        =   17
         Text            =   "*"
         Top             =   960
         Width           =   1995
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   15
         Top             =   1950
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.CommandButton cmdViewJobBill 
         Caption         =   "ดูข้อมูลใบ Menifest"
         Height          =   435
         Left            =   60
         TabIndex        =   3
         Top             =   4020
         Width           =   2205
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   60
         TabIndex        =   4
         Top             =   4470
         Width           =   2205
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   60
         TabIndex        =   2
         Top             =   3570
         Width           =   2205
      End
      Begin VB.ComboBox cboJobStatus 
         Height          =   315
         ItemData        =   "frCarWorkingQue.frx":0000
         Left            =   180
         List            =   "frCarWorkingQue.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1320
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
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   16
         Top             =   2610
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   2220
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor, คนขับ, ทะเบียนรถ:"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   720
         Width           =   1965
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   390
         TabIndex        =   14
         Top             =   5970
         Width           =   480
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปิดงานแล้ว"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   390
         TabIndex        =   13
         Top             =   5670
         Width           =   780
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "กำลังใช้งาน"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   390
         TabIndex        =   12
         Top             =   5370
         Width           =   810
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รอดำเนินการ"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   390
         TabIndex        =   11
         Top             =   5070
         Width           =   945
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   120
         Top             =   5940
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   120
         Top             =   5640
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   120
         Top             =   5340
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   120
         Top             =   5040
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   1710
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลข Manifest, ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   7
         Top             =   120
         Width           =   1650
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2370
      TabIndex        =   5
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
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รถที่ใช้ขน"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vendor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "คนขับ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "วันที่ขน"
         Object.Width           =   1984
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Menifest No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ลูกค้า"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
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
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อรุ่น,ลาย หรือยี่ห้อ :"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   4950
      Width           =   1440
   End
End
Attribute VB_Name = "frmCarWorkingQue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCondition_ As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    SearchQue
    If lvResult.ListItems.Count = 0 Then MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
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
        .AddItem "== ทุกสถานะ =="
        .AddItem "งานใหม่"
        .AddItem "พิมพ์ Menifest แล้ว"
        .AddItem "ปิดงานแล้ว"
        .AddItem "ยกเลิก"
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
            .Width = Me.Width - Picture1.Width - 100
        End If
        .Height = Me.Height - .Top - 410
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VisibleImage
End Sub

Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvResult.SortKey = ColumnHeader.Index - 1
    If lvResult.SortOrder = lvwAscending Then
        lvResult.SortOrder = lvwDescending
    Else
        lvResult.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvResult_DblClick()
On Error Resume Next
    If lvResult.SelectedItem Is Nothing Then
        MsgBox "กรุณาเลือกรถที่ต้องการปิด", vbExclamation
        Exit Sub
    End If
    Load frmCloseCar
    With frmCloseCar
        .lbCarDesc.Caption = lvResult.SelectedItem.Text
        .Show 1
    End With
End Sub

Private Sub Timer1_Timer()
    'SearchQue
End Sub

Private Sub SearchQue()
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim tmpList As ListItem
Dim x As Long, Y As Long
Dim sqlStr As String, strCond As String, StatusDesc As String
Dim tmpArr
Dim curColor
Dim subStr As String
    sqlStr = "Select MenifestID,DocNo,MenifestNo,CustomerName,SubTypeDesc,TruckTypeDesc,TransportName1,TransportName2,WorkDate,isClosed,isCanceled,isPrinted,DriverName,CarRegisID From vw_MenifestQue "
    strCond = " WHERE (WorkDate>='" & dtFrom.ValueYMD & "' And WorkDate<='" & dtTo.ValueYMD & "') "
    If Trim(txtKey.Text) <> "" And Replace(Trim(txtKey.Text), "*", "") <> "" Then
        strCond = strCond & " And (MenifestNo Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or CustomerName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')"
    End If
    If Trim(txtKey2.Text) <> "" And Replace(Trim(txtKey2.Text), "*", "") <> "" Then
        strCond = strCond & " And (TransportName1 Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or TransportName2 Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or DriverName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or CarRegisID Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')"
    End If
    Select Case cboJobStatus.ListIndex
        Case 1 'งานใหม่
            subStr = " ((isPrinted<>'Y' Or isPrinted is null) and (isCanceled<>'Y' Or isCanceled is null)) and (isClosed<>'Y' Or isClosed is null)) "
        Case 2 'พิมพ์ Menifest แล้ว
            subStr = " (isPrinted='Y' and (isCanceled<>'Y' Or isCanceled is null) and (isClosed<>'Y' Or isClosed is null)) "
        Case 3 'ปิดงาน
            subStr = " (isClosed='Y' and (isCanceled<>'Y' Or isCanceled is null)) "
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
    curCondition_ = strCond
    Screen.MousePointer = 11
    Set tmpRec = GetRS(sqlStr & strCond & " ORDER BY MenifestNo", adOpenForwardOnly, adLockReadOnly)
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
        tmpArr = tmpRec.GetRows
        For x = 0 To UBound(tmpArr, 2)
            Set tmpList = lvResult.ListItems.Add(, "'" & tmpArr(0, x) & "'", "" & tmpArr(13, x) & ", " & Trim("" & tmpArr(5, x)) & ", " & Trim("" & tmpArr(4, x)))
            If Trim("" & tmpArr(11, x)) <> "Y" And Trim("" & tmpArr(10, x)) <> "Y" And Trim("" & tmpArr(9, x)) <> "Y" Then
                curColor = lbOpen.ForeColor
                StatusDesc = lbOpen.Caption
            ElseIf Trim("" & tmpArr(11, x)) = "Y" And Trim("" & tmpArr(10, x)) <> "Y" And Trim("" & tmpArr(9, x)) <> "Y" Then
                curColor = lbApproved.ForeColor
                StatusDesc = lbApproved.Caption
            ElseIf Trim("" & tmpArr(9, x)) = "Y" And Trim("" & tmpArr(10, x)) <> "Y" Then
                curColor = lbClosed.ForeColor
                StatusDesc = lbClosed.Caption
            ElseIf Trim("" & tmpArr(10, x)) = "Y" Then
                curColor = lbCancel.ForeColor
                StatusDesc = lbCancel.Caption
            End If
            tmpList.ForeColor = curColor
            tmpList.SubItems(1) = Trim("" & tmpArr(7, x))
            tmpList.SubItems(2) = Trim("" & tmpArr(12, x))
            tmpList.SubItems(3) = FormatYMD_to_DMY(Trim("" & tmpArr(8, x)), "/", "/")
            tmpList.SubItems(4) = Trim("" & tmpArr(2, x))
            tmpList.SubItems(5) = Trim("" & tmpArr(3, x))
            tmpList.SubItems(6) = StatusDesc
            For Y = 1 To tmpList.ListSubItems.Count
                tmpList.ListSubItems(Y).ForeColor = curColor
            Next
            Set tmpList = Nothing
            'DoEvents
        Next
        Erase tmpArr
    End If
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

