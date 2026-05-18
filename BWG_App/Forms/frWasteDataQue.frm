VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmWasteDataQue 
   BackColor       =   &H00FFFFFF&
   Caption         =   "คิวของเสียรอการตรวจสอบ (Envi) - F150"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4290
      Top             =   5220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWasteDataQue.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWasteDataQue.frx":0296
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWasteDataQue.frx":02E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWasteDataQue.frx":05B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWasteDataQue.frx":087F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWasteDataQue.frx":0B4E
            Key             =   ""
         EndProperty
      EndProperty
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
      TabIndex        =   6
      Top             =   0
      Width           =   2385
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "เฉพาะรายการที่ยังไม่มีวันหมดอายุใบอนุญาต"
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   210
         TabIndex        =   18
         Top             =   3120
         Width           =   1905
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "เฉพาะรายการที่ยังไม่ออกเลข กอ."
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   210
         TabIndex        =   17
         Top             =   2520
         Width           =   1815
      End
      Begin WasteManagment.ctlDate dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   15
         Top             =   1380
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.CommandButton cmdViewWasteData 
         Caption         =   "ดูข้อมูลกากอุตสาหกรรม"
         Height          =   435
         Left            =   60
         TabIndex        =   3
         Top             =   4440
         Width           =   2205
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   60
         TabIndex        =   4
         Top             =   4890
         Width           =   2205
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   60
         TabIndex        =   2
         Top             =   3990
         Width           =   2205
      End
      Begin VB.ComboBox cboJobStatus 
         Height          =   315
         ItemData        =   "frWasteDataQue.frx":0DCB
         Left            =   180
         List            =   "frWasteDataQue.frx":0DCD
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
      Begin WasteManagment.ctlDate dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   16
         Top             =   2040
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   390
         TabIndex        =   14
         Top             =   6480
         Width           =   480
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตรวจสอบแล้ว, ไม่รับกำจัด"
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   390
         TabIndex        =   13
         Top             =   6180
         Width           =   1815
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตรวจสอบแล้ว, รับกำจัด"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   390
         TabIndex        =   12
         Top             =   5880
         Width           =   1620
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รอการตรวจสอบ"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   390
         TabIndex        =   11
         Top             =   5580
         Width           =   1110
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   120
         Top             =   6450
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF00FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   120
         Top             =   6150
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   120
         Top             =   5850
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   120
         Top             =   5550
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
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตั้งแต่วันที่"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   9
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อของเสีย, ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   7
         Top             =   120
         Width           =   1470
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
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "วันที่รับเข้า"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อของเสีย"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ลูกค้า"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Waste Code"
         Object.Width           =   2618
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Non/Haz?"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "บำบัด?"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ออกเลข กอ.?"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ใบอนุญาตหมดอายุ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "วิธีกำจัด"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ผล LAB"
         Object.Width           =   1411
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
Attribute VB_Name = "frmWasteDataQue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colDealerGroup_ As Collection
Dim colCurShowRoom_ As Collection
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
        If Me.ActiveControl.Container.Name = txtKey.Container.Name Or Me.ActiveControl.Container.Name = cboJobStatus.Container.Name Then
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
    CanAccess = False: CanUpdate = False
    SetPermission Me, CanAccess, CanUpdate
    If CanAccess = False Then
        MsgBox "ท่านไม่มีสิทธิใช้ระบบนี้", vbExclamation, "Warning"
        Unload Me
    End If
    
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
        .AddItem "รอการตรวจสอบ"
        .AddItem "ตรวจสอบแล้ว"
        .AddItem "บำบัดได้"
        .AddItem "บำบัดไม่ได้"
        .AddItem "ยกเลิก"
        .ListIndex = 0
    End With
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
    'MsgBox ColumnHeader.Width
    lvResult.SortKey = ColumnHeader.Index - 1
    If lvResult.SortOrder = lvwAscending Then
        lvResult.SortOrder = lvwDescending
    Else
        lvResult.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvResult_DblClick()
On Error Resume Next
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    Set tmpCol = tmpS.CustomerSearch("CustomerID in (Select CustomerID from tbWasteDataCR WHERE WasteDataID=" & lvResult.SelectedItem.Key & ")")
    If tmpCol.Count > 0 Then
        Load frmWasteData
        With frmWasteData
            .SetCallerForm Me, tmpCol(1)
            .SetDataWasteData tmpCol(1).WasteItems(lvResult.SelectedItem.Key)
            .Show 1
        End With
    End If
    Set tmpCol = Nothing
    Set tmpS = Nothing
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
    Screen.MousePointer = 11
    sqlStr = "Select WasteDataID,CreatedDate,WasteNo,WasteName,CompanyName,EUCode1,EUCode2,EUCode3,isApproveByLab,WasteType,isTreatAble,isCancel,LicenseExpDate,KorOrNumber,TreatmentDesc,isLabResultNeeded From vw_WasteDataCRSearch "
    strCond = " WHERE (CreatedDate>='" & dtFrom.ValueYMD & "' And CreatedDate<='" & dtTo.ValueYMD & "') "
    If Trim(txtKey.Text) <> "" And Replace(Trim(txtKey.Text), "*", "") <> "" Then
        strCond = strCond & " And (WasteName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or CompanyName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')"
    End If
    Select Case cboJobStatus.ListIndex
        Case 1 'รอการตรวจสอบ
            subStr = " ((isTreatAble='' Or isTreatAble is null) And (isCancel<>'Y' Or isCancel is null)) "
        Case 2 'ตรวจสอบแล้ว
            subStr = " ((isTreatAble='Y' Or isTreatAble='N') And (isCancel<>'Y' Or isCancel is null)) "
        Case 3 'บำบัดได้
            subStr = " ((isTreatAble='Y') And (isCancel<>'Y' Or isCancel is null)) "
        Case 4 'บำบัดไม่ได้
            subStr = " ((isTreatAble='N') And (isCancel<>'Y' Or isCancel is null)) "
        Case 5 'ยกเลิก
            subStr = " (isCancel='Y') "
    End Select
    If cboJobStatus.ListIndex > 0 Then
        If InStr(1, UCase(strCond), "WHERE") > 0 Then
            strCond = strCond & " And " & subStr
        Else
            strCond = strCond & " WHERE " & subStr
        End If
    End If
    If Check1.Value = 1 Then
        If InStr(1, UCase(strCond), "WHERE") > 0 Then
            strCond = strCond & " And KorOrNumber='' "
        Else
            strCond = strCond & " WHERE KorOrNumber='' "
        End If
    End If
    If Check2.Value = 1 Then
        If InStr(1, UCase(strCond), "WHERE") > 0 Then
            strCond = strCond & " And (LicenseExpDate='' Or LicenseExpDate is null) "
        Else
            strCond = strCond & " WHERE (LicenseExpDate='' Or LicenseExpDate is null) "
        End If
    End If
    curCondition_ = strCond
    
    sqlStr = sqlStr & strCond
    sqlStr = SelectWithCompanyID(sqlStr, False)
    sqlStr = sqlStr & " ORDER BY WasteNo"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then
        Screen.MousePointer = 0
        frmMain.Stb1.Panels(1).Text = "พบข้อมูล 0 รายการ"
        Exit Sub
    End If
    lvResult.ListItems.Clear
    If tmpRec.EOF = False Then
        tmpArr = tmpRec.GetRows
        For x = 0 To UBound(tmpArr, 2)
            Set tmpList = lvResult.ListItems.Add(, "'" & tmpArr(0, x) & "'", FormatYMD_to_DMY("" & tmpArr(1, x), "/", "/"))
            If Trim("" & tmpArr(10, x)) = "" And Trim("" & tmpArr(11, x)) <> "Y" Then
                curColor = lbOpen.ForeColor
                StatusDesc = "รอการตรวจสอบ"
            ElseIf Trim("" & tmpArr(10, x)) = "Y" And Trim("" & tmpArr(11, x)) <> "Y" Then
                curColor = lbApproved.ForeColor
                StatusDesc = "บำบัดได้"
            ElseIf Trim("" & tmpArr(10, x)) = "N" And Trim("" & tmpArr(11, x)) <> "Y" Then
                curColor = lbClosed.ForeColor
                StatusDesc = "บำบัดไม่ได้"
            ElseIf Trim("" & tmpArr(11, x)) = "Y" Then
                curColor = lbCancel.ForeColor
                StatusDesc = "ยกเลิก"
            End If
            tmpList.ForeColor = curColor
            tmpList.SubItems(1) = Trim("" & tmpArr(3, x))
            tmpList.SubItems(2) = Trim("" & tmpArr(4, x))
            tmpList.SubItems(3) = Trim("" & tmpArr(5, x)) & "-" & Trim("" & tmpArr(6, x)) & "-" & Trim("" & tmpArr(7, x))
            tmpList.SubItems(4) = Trim("" & tmpArr(9, x))
            tmpList.SubItems(5) = StatusDesc
            tmpList.SubItems(6) = Trim("" & tmpArr(13, x))
            If Trim("" & tmpArr(12, x)) <> "" Then
                tmpList.SubItems(7) = FormatYMD_to_DMY(Trim("" & tmpArr(12, x)), "/", "/")
            Else
                tmpList.SubItems(7) = " "
            End If
            tmpList.SubItems(8) = Trim("" & tmpArr(14, x))
            If Trim("" & tmpArr(15, x)) = "Y" Then
                tmpList.SubItems(9) = "ต้องการ"
            Else
                tmpList.SubItems(9) = "ไม่ต้องการ"
            End If
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
