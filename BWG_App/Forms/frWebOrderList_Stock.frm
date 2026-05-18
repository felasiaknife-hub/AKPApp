VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frWebOrderList_Stock 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Web order job list"
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
            Picture         =   "frWebOrderList_Stock.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWebOrderList_Stock.frx":0296
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWebOrderList_Stock.frx":02E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWebOrderList_Stock.frx":05B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWebOrderList_Stock.frx":087F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frWebOrderList_Stock.frx":0B4E
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
      TabIndex        =   8
      Top             =   0
      Width           =   2385
      Begin prjBWG.ctlDate dtFrom 
         Height          =   315
         Left            =   210
         TabIndex        =   19
         Top             =   2100
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.ComboBox cboDealerGroup 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   2025
      End
      Begin VB.ComboBox cboShowRoom 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1440
         Width           =   2025
      End
      Begin VB.CommandButton cmdViewJobBill 
         Caption         =   "ดูข้อมูลใบ Web Order"
         Height          =   435
         Left            =   60
         TabIndex        =   5
         Top             =   4440
         Width           =   2205
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "ปิด"
         Height          =   435
         Left            =   60
         TabIndex        =   6
         Top             =   4890
         Width           =   2205
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "ค้นหา"
         Height          =   435
         Left            =   60
         TabIndex        =   4
         Top             =   3990
         Width           =   2205
      End
      Begin VB.ComboBox cboJobStatus 
         Height          =   315
         ItemData        =   "frWebOrderList_Stock.frx":0DCB
         Left            =   180
         List            =   "frWebOrderList_Stock.frx":0DCD
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
      Begin prjBWG.ctlDate dtTo 
         Height          =   315
         Left            =   210
         TabIndex        =   20
         Top             =   2760
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BackColor       =   -2147483633
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "มีรายการสินค้าแจ้งยกเลิก"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   390
         TabIndex        =   18
         Top             =   7170
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   90
         Picture         =   "frWebOrderList_Stock.frx":0DCF
         Top             =   7140
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lbCancel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ยกเลิก"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   390
         TabIndex        =   17
         Top             =   6780
         Width           =   480
      End
      Begin VB.Label lbRequest 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "แจ้งเปลี่ยนแปลง,ยกเลิก"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   390
         TabIndex        =   16
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Label lbClosed 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Closed"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   390
         TabIndex        =   15
         Top             =   6180
         Width           =   480
      End
      Begin VB.Label lbApproved 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Approved"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   390
         TabIndex        =   14
         Top             =   5880
         Width           =   690
      End
      Begin VB.Label lbOpen 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   390
         TabIndex        =   13
         Top             =   5580
         Width           =   390
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   4
         Left            =   120
         Top             =   6750
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   3
         Left            =   120
         Top             =   6450
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   120
         Top             =   6150
         Width           =   195
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H0000FF00&
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
         TabIndex        =   12
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ตั้งแต่วันที่"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   11
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "เลข Manifest, ชื่อลูกค้า :"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   9
         Top             =   120
         Width           =   1650
      End
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2595
      Left            =   2370
      TabIndex        =   7
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
         Text            =   "Manifest"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ลูกค้า"
         Object.Width           =   2036
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ของเสีย"
         Object.Width           =   2884
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "รถที่ใช้ขน"
         Object.Width           =   2618
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "วันที่ขน"
         Object.Width           =   1984
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "วันที่รับ"
         Object.Width           =   1984
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
      TabIndex        =   10
      Top             =   4950
      Width           =   1440
   End
End
Attribute VB_Name = "frWebOrderList_Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim colDealerGroup_ As Collection
Dim colCurShowRoom_ As Collection
Dim curCondition_ As String

Private Sub cboDealerGroup_Click()
'Dim x%
'Dim tmpS As New BSPSearchManager
'    cboShowRoom.Clear
'    If colDealerGroup_ Is Nothing Then Exit Sub
'    If cboDealerGroup.ListIndex = -1 Then Exit Sub
'    Set colCurShowRoom_ = tmpS.CompanySearch("CompanyGroupID='" & colDealerGroup_(cboDealerGroup.ListIndex + 1).ID & "'")
'    For x = 1 To colCurShowRoom_.Count
'        cboShowRoom.AddItem colCurShowRoom_(x).CompanyCode
'    Next
'    cboShowRoom.AddItem "== ทุกโชว์รูม ==", 0
'    cboShowRoom.ListIndex = 0
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSearch_Click()
    SearchQue
    If lvResult.ListItems.Count = 0 Then MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
End Sub

Private Sub cmdViewJobBill_Click()
'    If lvResult.SelectedItem Is Nothing Then
'        MsgBox "กรุณาเลือกข้อมูล Web Order ที่ต้องการ", vbExclamation
'        Exit Sub
'    End If
'    Load frWebOrderItem
'    With frWebOrderItem
'        .WebOrderID = Replace(lvResult.SelectedItem.Key, "'", "")
'        .Show 1
'    End With
End Sub

Private Sub Command1_Click()
Dim tmpRec As New ADODB.Recordset
    tmpRec.Open "Select * from tbProductItem", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    Do Until tmpRec.EOF
        If IsNumeric("" & tmpRec("NetSalePrice")) = True Then
            tmpRec("RetailPrice") = RoundNumber(tmpRec("NetSalePrice") * 1.25)
            tmpRec.Update
        End If
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    MsgBox "Done"
End Sub

Private Sub Command2_Click()
Dim tmpRec As New ADODB.Recordset
Dim maxID As Long
    tmpRec.Open "select max(itembarcode) as maxID from tbproductitem", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    maxID = CLng(Replace(Trim("" & tmpRec("MaxID")), "PT", "")) + 1
    tmpRec.Close
    tmpRec.Open "select * from tbProductitem where (itembarcode is null or itembarcode='') order by subgroupid", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    Do Until tmpRec.EOF = True
        tmpRec("ItemBarcode") = "PT" & Format(maxID, "0000000")
        tmpRec.MoveNext
        maxID = maxID + 1
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    DBManager.GetDBConnection.Execute "Insert into tbRunProductCode (RunID,YearAdd,RunNumber) Values ('" & GetGUID & "','2549'," & (maxID - 1) & ")"
    MsgBox "Completed"
End Sub

'Private Sub Command3_Click()
'Dim tmpRec1 As New ADODB.Recordset
'Dim tmpRec2 As New ADODB.Recordset
'Dim tmpCon As ADODB.Connection
'Dim tmpID As String
'Dim tmpRowID As Long
'Dim tmpCol As New Collection
'Dim tmpAA As New BSPRecDetail
'    Set tmpCon = DBManager.GetDBConnection
'    tmpRec1.Open "select * from serialitem1 where weborderid<>'' and weborderid is not null", tmpCon, adOpenStatic, adLockReadOnly
'    Do Until tmpRec1.EOF
'        tmpRec2.Open "Select * from tbTakeStockHeader where WebOrderID='" & tmpRec1("WebOrderID") & "'", tmpCon, adOpenDynamic, adLockOptimistic
'        If tmpRec2.EOF Then
'            tmpID = GetGUID
'            tmpRec2.AddNew
'            tmpRec2("TakeStockID") = Trim(tmpID)
'            tmpRec2("WebOrderID") = tmpRec1("WebOrderID")
'            tmpRec2("TakeStockDate") = tmpRec1("TakeOutDate")
'            tmpRec2("TakeStaffID") = tmpRec1("TakeStaffID")
'            tmpRec2("StockStaffID") = tmpRec1("TakeStaffID")
'            tmpRec2("isCancel") = "N"
'            tmpRec2("CompanyID") = "CC20051213223211"
'            tmpRec2("IssueDStockID") = "9C9DD45F98CC4803B313482EB105E019"
'            tmpRec2.Update
'        Else
'            tmpID = tmpRec2("TakeStockID")
'        End If
'        tmpRec2.Close
'        tmpRec2.Open "Select * from tbWebOrderDetail Where WebOrderID='" & tmpRec1("WebOrderID") & "' And ProductItemID='" & tmpRec1("ProductItemID") & "'", tmpCon, adOpenDynamic, adLockOptimistic
'        If tmpRec2.EOF = False Then
'            If isExist(tmpCol, "'" & tmpRec2("ProductItemID") & "_" & tmpRec2("WebOrderID") & "'") Then
'                Set tmpAA = tmpCol("'" & tmpRec2("ProductItemID") & "_" & tmpRec2("WebOrderID") & "'")
'                tmpRowID = tmpAA.RecQty + 1
'                tmpAA.RecQty = tmpRowID
'                Set tmpAA = Nothing
'            Else
'                Set tmpAA = New BSPRecDetail
'                tmpAA.RecQty = 1
'                tmpRowID = 1
'                tmpCol.Add tmpAA, "'" & tmpRec2("ProductItemID") & "_" & tmpRec2("WebOrderID") & "'"
'                Set tmpAA = Nothing
'            End If
'            tmpCon.Execute "Insert into tbTakeStockDetail (TakeStockDetID,TakeStockID,WebOrderDetID,SerialID,isCuttingFilm,RowID,TakeQty) Values ('" & GetGUID & "','" & tmpID & "','" & tmpRec2("WebOrderDetID") & "','" & tmpRec1("SerialID") & "','N'," & tmpRowID & ",1)"
'            If Trim("" & tmpRec1("TakeSTaffID")) <> "" Then
'                tmpCon.Execute "Update tbSerialItem Set TakeStockID='" & tmpID & "',TakeStaffID='" & tmpRec1("TakeSTaffID") & "',TakeOutDate='" & tmpRec1("TakeOUtDate") & "' Where SerialID='" & tmpRec1("SerialID") & "'"
'            Else
'                tmpCon.Execute "Update tbSerialItem Set TakeStockID='" & tmpID & "',TakeStaffID='002023',TakeOutDate='" & tmpRec1("TakeOUtDate") & "' Where SerialID='" & tmpRec1("SerialID") & "'"
'            End If
'        End If
'        tmpRec2.Close
'        tmpRec1.MoveNext
'    Loop
'    tmpRec1.Close
'    Set tmpRec1 = Nothing
'    Set tmpRec2 = Nothing
'    Set tmpCon = Nothing
'    MsgBox "Updated"
'End Sub
'
'Private Sub Command4_Click()
'Dim tmpRec As New ADODB.Recordset
'Dim recCount As Long, iCount As Long
'    tmpRec.Open "SELECT     dbo.tbWebOrderDetail.WebOrderDetID, dbo.tbDealerPrice.DealerPrice " & _
'                            "FROM         dbo.tbWebOrderDetail INNER JOIN " & _
'                      "dbo.tbWebOrderHeader ON dbo.tbWebOrderDetail.WebOrderID = dbo.tbWebOrderHeader.WebOrderID INNER JOIN " & _
'                      "dbo.tbDealerPrice ON dbo.tbWebOrderDetail.ProductItemID = dbo.tbDealerPrice.ProductItemID INNER JOIN " & _
'                      "dbo.tbCompanyConfig ON dbo.tbDealerPrice.DealerGroupID = dbo.tbCompanyConfig.CompanyGroupID AND " & _
'                      "dbo.tbWebOrderHeader.DealerCompanyID = dbo.tbCompanyConfig.CompanyID where tbWebOrderDetail.DealerPrice=0", DBManager.GetDBConnection, adOpenStatic, adLockReadOnly
'    With DBManager.GetDBConnection
'        Do Until tmpRec.EOF
'            .Execute "Update tbWebOrderDetail Set DealerPrice=" & tmpRec("DealerPrice") & " Where WebOrderDetID='" & tmpRec("WebOrderDetID") & "'", iCount
'            recCount = recCount + iCount
'            tmpRec.MoveNext
'        Loop
'    End With
'    tmpRec.Close
'    Set tmpRec = Nothing
'    MsgBox "Done " & recCount & " records were updated"
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
'        If Me.ActiveControl.Container.Name = txtKey.Container.Name Or Me.ActiveControl.Container.Name = txtEngineNo.Container.Name Then
'            If Not TypeOf Me.ActiveControl Is CommandButton Then
'                cmdSearch_Click
'            End If
'        End If
'    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
Dim tmpRec As New ADODB.Recordset
Dim backDate As String
Dim tmpY As Integer, tmpM As Integer, tmpD As Integer
Dim tmpArr
'    tmpRec.Open "select backdate=convert(varchar(10),getdate()-7,111)", DBManager.GetDBConnection, adOpenStatic, adLockReadOnly
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
    With cboJobStatus
        .AddItem "== ทุกสถานะ =="
        .AddItem "Open"
        .AddItem "Approved"
        .AddItem "Closed"
        .AddItem "แจ้งยกเลิก"
        .AddItem "ยกเลิก"
        .AddItem "มีสินค้าแจ้งยกเลิก"
        .ListIndex = 0
    End With
    'PopulateDealerGroup
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
    'MsgBox ColumnHeader.Width
    lvResult.SortKey = ColumnHeader.Index - 1
    If lvResult.SortOrder = lvwAscending Then
        lvResult.SortOrder = lvwDescending
    Else
        lvResult.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvResult_DblClick()
'    If lvResult.SelectedItem Is Nothing Then Exit Sub
'    Load frJobOpen
'    frJobOpen.WebOrderID = Replace(lvResult.SelectedItem.Key, "'", "")
'    frJobOpen.Show 1
End Sub

Private Sub Timer1_Timer()
    SearchQue
End Sub

Private Sub SearchQue()
'Dim tmpRec As New ADODB.Recordset
'Dim tmpList As ListItem
'Dim x As Long, y As Long
'Dim sqlStr As String, strCond As String, statusDesc As String
'Dim tmpArr
'Dim curColor
'Dim subStr As String
'    sqlStr = "Select WebOrderID,DealerCompanyID,DealerName,StaffFName,StaffLName,WebJobNo,CarEngineNo,ChasisNo,BrandName,SeriesName,ModelName,CustomerName,OrderOpenDate,SellTime,CarRecDate,isApproved,isClose,CancelStatus,DealerCode,isJobOpen,CountRequest,isCutLosted From vw_WebOrderHeaderSearch "
'    If Trim(txtJobNo.Text) = "" Then
'        strCond = " Where (OrderOpenDate>='" & dtFrom.ValueYMD & "' And OrderOpenDate<='" & dtTo.ValueYMD & "') "
'        If Trim(txtKey.Text) <> "" And Replace(Trim(txtKey.Text), "*", "") <> "" Then
'            strCond = strCond & " And (StaffFName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%' Or StaffLName Like '%" & Replace(Trim(txtKey.Text), "*", "%") & "%')"
'        End If
'        If Trim(txtEngineNo.Text) <> "" Then
'            If InStr(1, LCase(strCond), "where") > 0 Then
'                strCond = strCond & " And CarEngineNo='" & Trim(txtEngineNo.Text) & "'"
'            Else
'                strCond = strCond & " Where CarEngineNo='" & Trim(txtEngineNo.Text) & "'"
'            End If
'        End If
'        If Trim(txtChasisNo.Text) <> "" Then
'            If InStr(1, LCase(strCond), "where") > 0 Then
'                strCond = strCond & " And ChasisNo='" & Trim(txtChasisNo.Text) & "'"
'            Else
'                strCond = strCond & " Where ChasisNo='" & Trim(txtChasisNo.Text) & "'"
'            End If
'        End If
'        Select Case cboJobStatus.ListIndex
'            Case 1 'Open
'                subStr = " (isApproved<>'Y'and isClose<>'Y' and (CancelStatus='' Or CancelStatus is null)) "
'            Case 2 'Approve
'                subStr = " (isApproved='Y'and isClose<>'Y' and (CancelStatus='' Or CancelStatus is null)) "
'            Case 3 'Closed
'                subStr = " (isClose='Y' and (CancelStatus='' Or CancelStatus is null)) "
'            Case 4 'Cancel Request
'                subStr = " (isClose<>'Y' and CancelStatus='R') "
'            Case 5 'Cancel
'                subStr = " (CancelStatus='C') "
'            Case 6
'                subStr = " (CountRequest>0) "
'        End Select
'        If cboJobStatus.ListIndex > 0 Then
'            If InStr(1, LCase(strCond), "where") > 0 Then
'                strCond = strCond & " And " & subStr
'            Else
'                strCond = strCond & " Where " & subStr
'            End If
'        End If
        
'        strCond = strCond & " And DealerGroupID='" & colDealerGroup_(cboDealerGroup.ListIndex + 1).ID & "'"
'        If cboShowRoom.ListIndex > 0 Then
'            strCond = strCond & " And DealerCompanyID='" & colCurShowRoom_(cboShowRoom.ListIndex).ID & "'"
'        End If
'    Else
'        strCond = strCond & " Where WebJobNo='" & Trim(txtJobNo.Text) & "'"
'    End If
'    curCondition_ = strCond
'    Screen.MousePointer = 11
'    tmpRec.Open sqlStr & strCond & " Order By WebJobNo Desc", DBManager.GetDBConnection, adOpenStatic, adLockReadOnly
'    lvResult.ListItems.Clear
'    If tmpRec.EOF = False Then
'        tmpArr = tmpRec.GetRows
'        For x = 0 To UBound(tmpArr, 2)
'            Set tmpList = lvResult.ListItems.Add(, "'" & tmpArr(0, x) & "'", "" & tmpArr(5, x))
'            If Trim("" & tmpArr(15, x)) <> "Y" And Trim("" & tmpArr(16, x)) <> "Y" And Trim("" & tmpArr(17, x)) = "" Then
'                curColor = lbOpen.ForeColor
'                statusDesc = "Open"
'            ElseIf Trim("" & tmpArr(15, x)) = "Y" And Trim("" & tmpArr(16, x)) <> "Y" And Trim("" & tmpArr(17, x)) = "" Then
'                curColor = lbApproved.ForeColor
'                statusDesc = "Approved"
'            'ElseIf Trim("" & tmpArr(15, x)) = "Y" And Trim("" & tmpArr(16, x)) = "Y" And Trim("" & tmpArr(17, x)) = "" Then
'            ElseIf Trim("" & tmpArr(16, x)) = "Y" And Trim("" & tmpArr(17, x)) = "" Then
'                curColor = lbClosed.ForeColor
'                statusDesc = "Closed"
'            ElseIf Trim("" & tmpArr(16, x)) <> "Y" And Trim("" & tmpArr(17, x)) = "R" Then
'                curColor = lbRequest.ForeColor
'                statusDesc = "แจ้งเปลี่ยนแปลง, ยกเลิก"
'            ElseIf Trim("" & tmpArr(17, x)) = "C" Then
'                curColor = lbCancel.ForeColor
'                statusDesc = "ยกเลิก"
'            End If
'            tmpList.ForeColor = curColor
'            tmpList.SubItems(1) = Trim("" & tmpArr(18, x))
'            tmpList.SubItems(2) = tmpArr(3, x) & " " & tmpArr(4, x)
'            tmpList.SubItems(3) = Trim("" & tmpArr(6, x))
'            tmpList.SubItems(4) = Trim("" & tmpArr(7, x))
'            tmpList.SubItems(5) = Trim("" & tmpArr(8, x)) & " - " & Trim("" & tmpArr(9, x)) '& " - " & Trim("" & tmpArr(10, x))
'            tmpList.SubItems(6) = Trim("" & tmpArr(11, x))
'            tmpList.SubItems(7) = FormatYMD_to_DMY(Trim("" & tmpArr(12, x)), "/", "/")
'            tmpList.SubItems(8) = FormatYMD_to_DMY(Trim("" & tmpArr(14, x)), "/", "/")
'            tmpList.SubItems(9) = statusDesc
'            For y = 1 To tmpList.ListSubItems.Count
'                tmpList.ListSubItems(y).ForeColor = curColor
'            Next
'            If IsNumeric(Trim(tmpArr(20, x))) Then
'                If CLng(Trim(tmpArr(20, x))) > 0 Then
'                    tmpList.SmallIcon = 6
'                Else
'                    tmpList.SmallIcon = 2
'                End If
'            End If
'            If Trim(tmpArr(21, x)) = "Y" Then
'                tmpList.SmallIcon = 5
'            End If
'
'            Set tmpList = Nothing
'            'DoEvents
'        Next
'        Erase tmpArr
'    End If
'    Screen.MousePointer = 0
'    frMain.Stb1.Panels(1).Text = "พบข้อมูล " & Format(tmpRec.RecordCount, "#,##0") & " รายการ"
'    tmpRec.Close
'    Set tmpRec = Nothing
End Sub

'Private Sub PopulateDealerGroup()
'Dim tmpS As New BSPSearchManager
'Dim x%
'    'Set colDealerGroup_ = tmpS.CompanyGroupSearch("isDealer='Y'")
'    Set colDealerGroup_ = tmpS.CompanyGroupSearch()
'    cboDealerGroup.Clear
'    For x = 1 To colDealerGroup_.Count
'        cboDealerGroup.AddItem colDealerGroup_(x).GroupName
'    Next
'    If cboDealerGroup.ListCount > 0 Then cboDealerGroup.ListIndex = 0
'    Set tmpS = Nothing
'End Sub
'
