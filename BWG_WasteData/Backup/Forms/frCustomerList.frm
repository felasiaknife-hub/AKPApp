VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCustomerList 
   BackColor       =   &H00FFFFC0&
   Caption         =   "กากที่รอการตรวจรับ"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   ForeColor       =   &H00808080&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CRrpt 
      Left            =   2880
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
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
            Picture         =   "frCustomerList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCustomerList.frx":0296
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCustomerList.frx":02E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCustomerList.frx":05B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCustomerList.frx":087F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCustomerList.frx":0B4E
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
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   9420
      Left            =   0
      ScaleHeight     =   9390
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   0
      Width           =   2385
      Begin VB.ComboBox cboStatus 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1260
         Width           =   1995
      End
      Begin VB.TextBox txtKey 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Text            =   "*"
         Top             =   540
         Width           =   1995
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   60
         TabIndex        =   1
         Top             =   4020
         Width           =   2205
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   60
         TabIndex        =   0
         Top             =   3360
         Width           =   2205
      End
      Begin WasteDataForm.ctlDate dtTo 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   2700
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin WasteDataForm.ctlDate dtFrom 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   1980
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "สถานะ"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1020
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ชื่อของเสีย, ชื่อลูกค้า :"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตั้งแต่วันที่"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   495
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   120
         Top             =   5520
         Width           =   195
      End
      Begin VB.Label lblReject 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตีกลับ"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   420
         TabIndex        =   7
         Top             =   5520
         Width           =   405
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตรวจสอบแล้ว, รับเป็นลูกค้า"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   420
         TabIndex        =   6
         Top             =   5160
         Width           =   1920
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รอการตรวจสอบ"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   420
         TabIndex        =   5
         Top             =   4740
         Width           =   1155
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   120
         Top             =   5160
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   120
         Top             =   4740
         Width           =   195
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2340
      TabIndex        =   2
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "วันที่รับเข้า"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อลูกค้า"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ที่อยู่"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ผู้ติดต่อ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "สถานะ"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Sales"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ผู้ตรวจรับ"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "วันที่ตรวจรับ"
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อรุ่น,ลาย หรือยี่ห้อ :"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Top             =   4950
      Width           =   1440
   End
End
Attribute VB_Name = "frmCustomerList"
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
        If Me.ActiveControl.Container.Name = txtKey.Container.Name Or Me.ActiveControl.Container.Name = cboStatus.Container.Name Then
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
    
'    tmpRec.Open "select backdate=convert(varchar(10),getdate()-7,111)", DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
'    backDate = Trim("" & tmpRec("BackDate"))
'    tmpRec.Close
    If IsDate(backDate) Then
        tmpArr = Split(backDate, "/")
        If CInt(tmpArr(0)) < 2500 Then
            dtFrom.ValueYMD = Trim("" & (CInt(tmpArr(0) + 543))) & "/" & Trim("" & tmpArr(1)) & "/" & Trim("" & tmpArr(2))
        Else
            dtFrom.ValueYMD = backDate
        End If
    End If
    Set tmpRec = Nothing
    With cboStatus
        .AddItem "== ทุกสถานะ =="
        .AddItem "รอตรวจรับ"
        .AddItem "ตรวจรับแล้ว"
        .AddItem "ตีกลับ"
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
Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvResult.SortKey = ColumnHeader.Index - 1
    If lvResult.SortOrder = lvwAscending Then
        lvResult.SortOrder = lvwDescending
    Else
        lvResult.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvResult_DblClick()
On Error GoTo ErrD
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    If lvResult.ListItems.Count = 0 Then Exit Sub
    If lvResult.SelectedItem Is Nothing Then Exit Sub
    Set tmpCol = tmpS.CustomerSearch("CustomerID =" & lvResult.SelectedItem.Key)
    If tmpCol.Count > 0 Then
        Load frmCustomer
        With frmCustomer
            .CallLocation = "N"
            .SetDataCustomer tmpCol(1)
            .Show 1
        End With
    End If
    Set tmpCol = Nothing
    Set tmpS = Nothing
    Exit Sub
ErrD:
    Resume Next
End Sub

Private Sub Timer1_Timer()
    'SearchQue
End Sub

Private Sub SearchQue()
'On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim tmpList As ListItem
Dim x As Long, Y As Long, i As Integer
Dim sqlStr As String, sqlStr1 As String, sqlStr2 As String, strCond As String, statusDesc As String
Dim tmpArr
Dim curColor
Dim subStr As String, LastYear As String, CurYear As String
Dim TotalCount As Long, LastY As Integer, OneMonth As String
    LastY = Left(TodayDate, 4) - 1
    LastYear = LastY & Mid(TodayDate, 5)
    OneMonth = LastY & Format(Date + 30, "/mm/dd")
    CurYear = Format(Date + 30, "yyyy/mm/dd")
    
'    sqlStr = "SELECT     DISTINCT dbo.vw_Customer.CustomerID, dbo.vw_Customer.CreatedDate, dbo.vw_Customer.CompanyName, dbo.vw_Customer.Addr1," & vbCrLf
'    sqlStr = sqlStr & "                      dbo.vw_Customer.ContactPerson, dbo.vw_Customer.isApprove, dbo.vw_Customer.ApproveDate, dbo.vw_Customer.isNew," & vbCrLf
'    sqlStr = sqlStr & "                      dbo.vw_SaleDistrictSearch.StaffFName + '  ' + dbo.vw_SaleDistrictSearch.StaffLName AS SaleStaff, isRequest, RejectDesc," & vbCrLf
'    sqlStr = sqlStr & "                      dbo.vw_Customer.isApproveName, dbo.vw_Customer.Status" & vbCrLf
'    sqlStr = sqlStr & "FROM         dbo.vw_Customer LEFT OUTER JOIN" & vbCrLf
'    sqlStr = sqlStr & "                      dbo.vw_SaleDistrictSearch ON dbo.vw_Customer.UnderSaleStaffID = dbo.vw_SaleDistrictSearch.SaleStaffID INNER JOIN" & vbCrLf
'    sqlStr = sqlStr & "                      dbo.tbCompanyStaff ON dbo.vw_Customer.CreatedByStaffID = dbo.tbCompanyStaff.StaffID" & vbCrLf
'
    sqlStr = " Select CustomerID,CreatedDate + ' ' + Addtime as CreatedDate,CompanyName,Addr1," & vbCrLf
    sqlStr = sqlStr & " ContactPerson,isApprove,ApproveDate,isNew,StaffFName + ' ' + StaffLName +  ' '  + StaffCode  As SaleStaff, isRequest, RejectDesc," & vbCrLf
    sqlStr = sqlStr & " isApproveName,Status " & vbCrLf
    sqlStr = sqlStr & "From vw_CustomerAll " & vbCrLf
    If Trim(txtKey.Text) <> "" Or Replace(Trim(txtKey.Text), "*", "") <> "" Then
        strCond = "WHERE (CompanyName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')" & vbCrLf
    End If
    If InStr(1, LCase(strCond), "where") > 0 Then
        strCond = strCond & " And " & subStr
    Else
        strCond = strCond & " Where " & subStr
    End If
    strCond = strCond & " (CreatedDate >= '" & dtFrom.ValueYMD & "' And CreatedDate <= '" & dtTo.ValueYMD & "' ) " & vbCrLf
    Select Case cboStatus.ListIndex
        Case 1 'รอตรวจรับ
'            subStr = " (isNew = 'Y') and (isApprove = 'N') and (RejectDesc = '')"
            subStr = " (isNew = 'Y') and (isApprove = 'N') and (Status = 'R')"
        Case 2 'ตรวจรับแล้ว
'            subStr = " (isApprove = 'Y') and (isNew ='N')"
            subStr = " (isApprove = 'Y') and (isNew ='N') and (Status = 'N')"
        Case 3 'ตีกลับ
            subStr = " (isApprove = 'N') and (isNew ='Y') and (Status = 'J')"
    End Select
    If cboStatus.ListIndex > 0 Then
        If InStr(1, LCase(strCond), "where") > 0 Then
            strCond = strCond & " And " & subStr
        Else
            strCond = strCond & " Where " & subStr
        End If
    End If
    strCond = strCond & " And (Status <> 'Y')"
    
    If isAkp = True And isBWG = False And isPST = False Then
       strCond = strCond & " And (SDName like '%AP%')"
    ElseIf isAkp = False And isPST = True And isBWG = False Then
       strCond = strCond & " And (SDName like '%PST%')"
    End If
    
    sqlStr = sqlStr & strCond & vbCrLf
    sqlStr = sqlStr & "ORDER BY CompanyName"
    curCondition_ = strCond
    Screen.MousePointer = 11
    lvResult.ListItems.Clear
    Set tmpRec = New ADODB.Recordset
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF = False Then
        TotalCount = TotalCount + tmpRec.RecordCount
        tmpArr = tmpRec.GetRows
        For x = 0 To UBound(tmpArr, 2)
            Set tmpList = lvResult.ListItems.Add(, "'" & tmpArr(0, x) & "'", IIf(IsNull(tmpArr(2, x)), "", Format(tmpArr(1, x), "dd/mm/yyyy HH:mm")))
            If tmpArr(12, x) = "R" Then
'            If tmpArr(7, x) = "Y" And tmpArr(9, x) = "Y" Then
                statusDesc = "รอตรวจรับ"
                curColor = lbOpen.ForeColor
            ElseIf Trim("" & tmpArr(12, x)) = "N" Then
'            ElseIf Trim("" & tmpArr(5, x)) = "Y" Then
                statusDesc = "ตรวจรับแล้ว"
                curColor = lbApproved.ForeColor
            ElseIf tmpArr(12, x) = "J" Then
'            ElseIf tmpArr(10, x) <> "" Then
                statusDesc = "ตีกลับ"
                curColor = lblReject.ForeColor
            ElseIf tmpArr(12, x) = "D" Then
'            ElseIf tmpArr(7, x) = "Y" And tmpArr(9, x) = "N" Then
                statusDesc = "ยังไม่ส่งข้อมูล"
                curColor = vbBlack
            ElseIf tmpArr(12, x) = "P" And Trim("" & tmpArr(5, x)) = "Y" Then
                statusDesc = "ตรวจรับแล้ว"
                curColor = lbApproved.ForeColor
            End If
            Dim Name As String
            If tmpArr(11, x) = "R" Then
                Name = "-"
            Else
                Name = tmpArr(11, x)
            End If
            tmpList.ForeColor = curColor
            tmpList.SubItems(1) = IIf(IsNull(tmpArr(2, x)), "", tmpArr(2, x))
            tmpList.SubItems(2) = IIf(IsNull(tmpArr(3, x)), "", tmpArr(3, x))
            tmpList.SubItems(3) = IIf(IsNull(tmpArr(4, x)), "", tmpArr(4, x))
            tmpList.SubItems(4) = statusDesc
            tmpList.SubItems(5) = IIf(IsNull(tmpArr(8, x)), "", tmpArr(8, x))
            tmpList.SubItems(6) = Name
            tmpList.SubItems(7) = IIf(IsNull(tmpArr(6, x)), "", tmpArr(6, x))
            For Y = 1 To tmpList.ListSubItems.Count
                tmpList.ListSubItems(Y).ForeColor = curColor
            Next
            Set tmpList = Nothing
        Next
        Erase tmpArr
    End If
    Screen.MousePointer = 0
    frmMain.Stb1.Panels(1).Text = "พบข้อมูล " & Format(TotalCount, "#,##0") & " รายการ"
    tmpRec.Close
    Set tmpRec = Nothing
    Exit Sub
'ErrD:
'    Screen.MousePointer = 0
'    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
'    Exit Sub
End Sub
