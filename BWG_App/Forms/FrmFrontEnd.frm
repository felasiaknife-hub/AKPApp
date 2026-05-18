VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B4BF0775-2BEE-4C33-B2CC-39B9BC2E4704}#4.0#0"; "DateCtl.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmFrontEnd 
   Caption         =   "ข้อมูลการสั่งงาน Front end"
   ClientHeight    =   9540
   ClientLeft      =   3780
   ClientTop       =   1950
   ClientWidth     =   13545
   LinkTopic       =   "Form1"
   ScaleHeight     =   9540
   ScaleWidth      =   13545
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9540
      Left            =   0
      ScaleHeight     =   9510
      ScaleWidth      =   2625
      TabIndex        =   0
      Top             =   0
      Width           =   2655
      Begin VB.OptionButton OptWorkDate 
         BackColor       =   &H8000000E&
         Caption         =   "วันที่ขน"
         Height          =   195
         Left            =   1350
         TabIndex        =   23
         Top             =   1350
         Width           =   1095
      End
      Begin VB.OptionButton OptOrder 
         BackColor       =   &H8000000E&
         Caption         =   "วันที่สั่งงาน"
         Height          =   195
         Left            =   225
         TabIndex        =   22
         Top             =   1350
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         Caption         =   "สถานะ"
         Height          =   1770
         Left            =   180
         TabIndex        =   16
         Top             =   2610
         Width           =   2310
         Begin VB.OptionButton OptComplete 
            Caption         =   "เสร็จสิ้น"
            Height          =   195
            Left            =   315
            TabIndex        =   21
            Top             =   1125
            Width           =   1140
         End
         Begin VB.OptionButton OptCancel 
            Caption         =   "ยกเลิก"
            Height          =   195
            Left            =   315
            TabIndex        =   20
            Top             =   1440
            Width           =   1140
         End
         Begin VB.OptionButton OptApprove 
            Caption         =   "ยืนยัน"
            Height          =   195
            Left            =   315
            TabIndex        =   19
            Top             =   810
            Width           =   1140
         End
         Begin VB.OptionButton OptWait 
            Caption         =   "รอยืนยัน"
            Height          =   195
            Left            =   315
            TabIndex        =   18
            Top             =   495
            Width           =   1140
         End
         Begin VB.OptionButton OptAll 
            Caption         =   "ทั้งหมด"
            Height          =   195
            Left            =   315
            TabIndex        =   17
            Top             =   225
            Width           =   1140
         End
      End
      Begin VB.TextBox txtSale 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   180
         TabIndex        =   3
         Top             =   4935
         Width           =   2295
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   180
         TabIndex        =   2
         Top             =   4440
         Width           =   2295
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Text            =   "*"
         Top             =   360
         Width           =   2295
      End
      Begin DateCtl.Date dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   5
         Top             =   1575
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin DateCtl.Date dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   6
         Top             =   2235
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label LblTotal 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   90
         TabIndex        =   15
         Top             =   5490
         Width           =   75
      End
      Begin VB.Line Line2 
         X1              =   180
         X2              =   2490
         Y1              =   4410
         Y2              =   4410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "พนักงานขาย :"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   450
         TabIndex        =   12
         Top             =   9120
         Width           =   480
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ปิดใบเสนอราคา"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   450
         TabIndex        =   11
         Top             =   8820
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "อนุมัติแล้ว"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   450
         TabIndex        =   10
         Top             =   8520
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "งานใหม่"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   450
         TabIndex        =   9
         Top             =   8220
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   180
         Top             =   9090
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00800080&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   180
         Top             =   8790
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   180
         Top             =   8490
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   180
         Top             =   8190
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ถึงวันที่"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   8
         Top             =   1995
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "รหัสหรือชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   7
         Top             =   120
         Width           =   1230
      End
   End
   Begin Crystal.CrystalReport rpt1 
      Left            =   4980
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2640
      TabIndex        =   14
      Top             =   0
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   4577
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImgSorted"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "รหัส"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "วันที่สั่งงาน"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "รหัสลูกค้า"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "บริษัท"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "วันที่สั่งขน"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ข้อความ"
         Object.Width           =   2618
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "จำนวนรถ"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "คนงาน"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ยกเลิก"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ผู้สั่ง"
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "EMail"
         Object.Width           =   3351
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "โทร"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "การลงออเดอร์"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "เสร็จสิ้น"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "ServiceOrderID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "UnderSaleID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgSorted 
      Left            =   6720
      Top             =   7920
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
            Picture         =   "FrmFrontEnd.frx":0000
            Key             =   "up"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFrontEnd.frx":059A
            Key             =   "down"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmFrontEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SaleStaff_ As BWGCompanyStaff
Dim curCustomer_ As BWGCustomer
Dim colSearchResult_ As Collection

Private Function SQLString() As String
    Dim strCond As String
    Dim criStr As String
    criStr = Replace(Replace(Trim(txtCustomer.Text), "*", "%"), "'", "")
'    strCond = "(QuotationNo Like '%" & criStr & "%' Or CustomerCode Like '%" & criStr & "%' Or CompanyName Like '%" & criStr & "%')"
Dim dtf As String
Dim dtt As String
Dim Yf As String
Dim Yt As String
Dim DateFrom As String
Dim DateTo As String
dtf = dtFrom.ValueYMD
dtt = dtTo.ValueYMD
Yf = Left(dtf, 4)
Yt = Left(dtt, 4)
DateFrom = Yf - 543 & "-" & Format(dtf, "MM-dd")
DateTo = Yt - 543 & "-" & Format(dtt, "MM-dd")
If OptOrder.Value = True Then
    strCond = "(orderdate BETWEEN '" & DateFrom & "' AND '" & DateTo & " " & "23:59:59.999" & "')"
ElseIf OptWorkDate.Value = True Then
    strCond = "(workdate BETWEEN '" & DateFrom & "' AND '" & DateTo & "')"
End If
    If Not curCustomer_ Is Nothing Then strCond = strCond & vbCrLf & " AND CustomerID = '" & curCustomer_.ID & "'"
    If Not SaleStaff_ Is Nothing Then strCond = strCond & vbCrLf & " AND (UnderSaleStaffID = '" & SaleStaff_.ID & "')"
    
    If OptAll.Value = True Then
        strCond = strCond
    ElseIf OptWait.Value = True Then
        strCond = strCond & vbCrLf & " AND (ConfirmOrder <> 'Y' or confirmorder is null) and (isCancel <> 'Y')"
    ElseIf OptApprove.Value = True Then
        strCond = strCond & vbCrLf & " AND (ConfirmOrder = 'Y')"
    ElseIf OptCancel.Value = True Then
        strCond = strCond & vbCrLf & " AND (isCancel = 'Y')"
    ElseIf OptComplete.Value = True Then
        strCond = strCond & vbCrLf & " AND (completed = 'Y')"
    End If
    SQLString = strCond
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    On Error GoTo ErrD
    Dim criStr As String
    Dim strCond As String
    Dim tmpRec As New ADODB.Recordset
    Dim x%, result%, y%, i%
    Dim tmpList As ListItem
    Dim tmpS As BWGSearchManager
    Dim tmpItem As BWGQuotationHeader
    Dim tmpQuoDet As BWGQuotationDetail
    Dim tmpColDet As Collection
    
    Screen.MousePointer = 11
    Set colSearchResult_ = Nothing
    Set colSearchResult_ = New Collection
    lvResult.ListItems.Clear
    Set tmpS = New BWGSearchManager
    strCond = SQLString
    
'    strCond = SelectWithCompanyID(strCond, True)
'    Set colSearchResult_ = tmpS.QuotationDetailSearch(strCond)
'    Set tmpS = Nothing
'    Set tmpItem = colSearchResult_(x)
'    Set tmpColDet = tmpRec!QuotationDetails
    sqlStr = "SELECT * FROM vw_Service_OrderHD" & vbCrLf
    sqlStr = sqlStr & "WHERE " & strCond & vbCrLf
    If OptOrder.Value = True Then
    sqlStr = sqlStr & "ORDER BY  orderdate "
    ElseIf OptWorkDate.Value = True Then
    sqlStr = sqlStr & "ORDER BY  workdate "
    End If
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    
'    If tmpRec.RecordCount = 0 Then Screen.MousePointer = 0: Exit Sub
        If tmpRec.RecordCount = 0 Then
            LblTotal.Caption = "พบข้อมูลทั้งหมด " & Format(tmpRec.RecordCount, "#,##0") & " รายการ"
            LblTotal.ForeColor = vbRed
            LblTotal.BackColor = vbBlue
        Else
            LblTotal.Caption = "พบข้อมูลทั้งหมด " & Format(tmpRec.RecordCount, "#,##0") & " รายการ"
            LblTotal.ForeColor = vbGreen
            LblTotal.BackColor = vbBlack
        End If
    Screen.MousePointer = 0
    DoEvents
    Screen.MousePointer = 11
    Do While Not tmpRec.EOF
        Dim ForeColor
        Dim sqlSelTruck As String
        Dim tmpSel As New ADODB.Recordset
        Dim CancelOrd As String
        
        
        
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpRec!CustomerID & "_" & tmpRec!OrderID & "'", tmpRec!OrderNo)
        tmpList.SubItems(1) = Format(tmpRec!orderdate, "dd/MM/yyyy")
        tmpList.SubItems(2) = tmpRec!CustomerCode
        tmpList.SubItems(3) = Trim(tmpRec!CompanyName)
        tmpList.SubItems(4) = Format(tmpRec!WorkDate, "dd/MM/yyyy")
        tmpList.SubItems(5) = tmpRec!ordermessage
        tmpList.SubItems(6) = tmpRec!TruckCount
        tmpList.SubItems(7) = tmpRec!WorkerPerCar
        tmpList.SubItems(9) = Trim(tmpRec!Name) & " " & Trim(tmpRec!LasteName)
        tmpList.SubItems(10) = tmpRec!Email
        tmpList.SubItems(11) = tmpRec!TelNo
        
        If tmpRec!ConfirmOrder = "Y" Then
            tmpList.SubItems(12) = "ยืนยันแล้ว"
        Else
            tmpList.SubItems(12) = "รอยืนยัน"
        End If
        If tmpRec!completed = "Y" Then
            tmpList.SubItems(13) = "เสร็จสิ้นแล้ว"
        Else
            tmpList.SubItems(13) = "ยังไม่เสร็จสิ้น"
        End If
        tmpList.SubItems(14) = tmpRec!OrderID
        tmpList.SubItems(15) = tmpRec!UnderSaleStaffID
        If tmpRec!isCancel = "Y" Then
            tmpList.SubItems(8) = "ยกเลิก"
            tmpList.ForeColor = vbRed
            ForeColor = vbRed
        Else
            tmpList.SubItems(8) = "ปกติ"
            tmpList.ForeColor = vbBlack
            ForeColor = vbBlack
        End If
        
        
            For i = 1 To tmpList.ListSubItems.Count
    '                    tmpList.ListSubItems(i).Bold = True
            tmpList.ListSubItems(i).ForeColor = ForeColor
        Next

        DoEvents
        Set tmpList = Nothing
        tmpRec.MoveNext
    Loop
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
    Set tmpRec = Nothing
    Screen.MousePointer = 0
    Exit Sub
    
ErrD:
    MsgBox err.Description, vbCritical
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    On Error GoTo ErrD
    Dim tmpRec As New ADODB.Recordset
    Dim backDate As String, tmpArr
    Dim sqlStr As String
    OptAll.Value = True
    OptOrder.Value = True
    Me.Icon = frmMain.Icon
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
    Call ClearHeaderIcons(ColumnHeader.Index)
    Select Case ColumnHeader.Index
        Case 2
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
On Error GoTo ErrD
Dim tmpS As New BWGSearchManager, tmpCol As Collection

    If lvResult.SelectedItem Is Nothing Then Exit Sub
    Load frmJobOpen
    With frmJobOpen
        Set tmpCol = tmpS.JobDataHeaderSearch(" Service_OrderID= '" & lvResult.SelectedItem.SubItems(14) & "'")
        If tmpCol.Count > 0 Then
            .SetDataJobOpen tmpCol(1)
            Set tmpCol = Nothing
            Set tmpS = Nothing
            .Show 1
        Else
            .txtCustomer.Text = lvResult.SelectedItem.SubItems(2)
            .dtWorkDate.ValueDMY = lvResult.SelectedItem.SubItems(4)
            .SetTxtCustomer
            .setConfirmOrderFrontEnd (lvResult.SelectedItem.SubItems(14))
            .PopulateDistrictFrontEnd (lvResult.SelectedItem.SubItems(15))
            .cmdAddItem_Click
            .Show 1
        End If
    End With
    Set tmpCol = Nothing
    Set tmpS = Nothing
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub

End Sub

Private Sub TxtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCustomer_Validate False
End Sub

Private Sub txtCustomer_Validate(Cancel As Boolean)
    If curCustomer_ Is Nothing Then Set curCustomer_ = New BWGCustomer
    If Trim(txtCustomer.Text) = "*" Or Trim(txtCustomer.Text) = "" Then Set curCustomer_ = Nothing: Exit Sub
    If Trim(txtCustomer.Text) = "" And Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
        Set curCustomer_ = Nothing
        Set curCustomer_ = New BWGCustomer
    Else
        If Trim(curCustomer_.CustomerName) <> Trim(txtCustomer.Text) Then
            Set curCustomer_ = Nothing
            Set curCustomer_ = New BWGCustomer
            Load frmCustomerSearch
            With frmCustomerSearch
                .CStatus = "N"
                .getCriteria (txtCustomer.Text)
                .GetCriteriaOptonal " AND CustGroupID <> '488F2E5C057F46869D1A56185B11D644'"
                .SetCallerForm Me
                .Show 1
            End With
            If curCustomer_.ID = "" Then txtCustomer.Text = ""
        End If
    End If
End Sub

Public Sub SetDataCustomer(selCust As BWGCustomer, Optional isAssigned As Boolean = True)
    If selCust Is Nothing Then
        txtCustomer.Text = ""
        Exit Sub
    End If
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCust
    With curCustomer_
        txtCustomer.Text = .CustomerName
    End With
End Sub

Private Sub txtSale_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtSale_Validate False
End Sub

Private Sub txtSale_Validate(Cancel As Boolean)
    If SaleStaff_ Is Nothing Then Set SaleStaff_ = New BWGCompanyStaff
    If Trim(txtSale.Text) = "" And Trim(SaleStaff_.StaffFName & " " & SaleStaff_.StaffLName) <> Trim(txtSale.Text) Then
        Set SaleStaff_ = Nothing
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
    Set SaleStaff_ = Nothing
    Set SaleStaff_ = selStaff
    If SaleStaff_ Is Nothing Then Exit Sub
    With SaleStaff_
        txtSale.Text = .StaffFName & " " & .StaffLName
    End With
End Sub

