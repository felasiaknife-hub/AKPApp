VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJobWasteDataAdd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Waste List for added"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12330
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   12330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkAll 
      BackColor       =   &H00FFFFFF&
      Caption         =   "แสดงใบเสนอราคาทั้งหมด"
      Height          =   285
      Left            =   270
      TabIndex        =   18
      Top             =   1845
      Width           =   2220
   End
   Begin VB.TextBox TxtD1 
      Height          =   330
      Left            =   10260
      TabIndex        =   13
      Top             =   1620
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ค้นหาข้อมูล"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   870
      Width           =   2595
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "อัปเดทรายการที่เลือก / ปิด"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1290
      Width           =   2595
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "ยกเลิก"
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1710
      Width           =   2595
   End
   Begin VB.TextBox txtKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2670
      TabIndex        =   0
      Text            =   "*"
      Top             =   1350
      Width           =   3645
   End
   Begin MSComctlLib.ListView lvResult 
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   2460
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Waste No"
         Object.Width           =   2752
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อของเสีย"
         Object.Width           =   3545
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "เลขที่ใบเสนอราคา"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ใบเสนอราคา BME"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ค่ากำจัด"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "หน่วย"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ส่วนลดค่าขนส่ง"
         Object.Width           =   1589
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "วันที่อนุญาต"
         Object.Width           =   2206
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "หมดอายุ"
         Object.Width           =   2206
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ประเภทรถ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "รวมคนงาน"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "ค่าคนงานเพิ่ม"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "ประเภทใบเสนอราคา"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "สถานที่กำจัด (Default)"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "ค่าขนส่ง"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvSelected 
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   5400
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   3413
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Waste No"
         Object.Width           =   2752
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ชื่อของเสีย"
         Object.Width           =   3545
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "เลขที่ใบเสนอราคา"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ค่ากำจัด"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "หน่วย"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ประเภทรถ"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "รวมคนงาน"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "ค่าคนงานเพิ่ม"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "สถานที่กำจัด"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ค่าขนส่ง"
         Object.Width           =   2540
      EndProperty
   End
   Begin VSFlex7LCtl.VSFlexGrid GrOrderID 
      Height          =   2445
      Left            =   8775
      TabIndex        =   17
      ToolTipText     =   "สำหรับแสดงข้อมูลเท่านั้น ไม่สามารถแก้ไขได้"
      Top             =   1125
      Visible         =   0   'False
      Width           =   3435
      _cx             =   6059
      _cy             =   4313
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421504
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   12
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label LblConfirm 
      Height          =   330
      Left            =   270
      TabIndex        =   16
      Top             =   900
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label4 
      Height          =   285
      Left            =   10440
      TabIndex        =   15
      Top             =   1305
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Label3 
      Height          =   240
      Left            =   11070
      TabIndex        =   14
      Top             =   990
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อของเสีย, Waste No, เลขที่ใบเสนอราคา"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   2670
      TabIndex        =   12
      Top             =   1710
      Width           =   2895
   End
   Begin VB.Label lbCustomer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2700
      TabIndex        =   11
      Top             =   990
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ชื่อลูกค้า"
      Height          =   195
      Index           =   0
      Left            =   2010
      TabIndex        =   10
      Top             =   990
      Width           =   570
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   -60
      X2              =   12420
      Y1              =   5130
      Y2              =   5130
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "รายการสินค้าที่เลือก"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   5160
      Width           =   12345
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   0
      X2              =   12360
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ผลการค้นหา"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   2220
      Width           =   12345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "คำที่ใช้ในการค้นหา"
      Height          =   195
      Index           =   7
      Left            =   1215
      TabIndex        =   7
      Top             =   1410
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ค้นหารายการกากของเสียเพื่อสั่งขน"
      BeginProperty Font 
         Name            =   "Angsana New"
         Size            =   18
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   90
      Width           =   3390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   7050
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   4
      Index           =   0
      X1              =   150
      X2              =   6960
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Image Image2 
      Height          =   795
      Left            =   0
      Picture         =   "frJobWasteDataAdd.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12420
   End
End
Attribute VB_Name = "frmJobWasteDataAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer
Dim curWasteGenNo_ As String
Dim curFactoryRegisNo_ As String
Dim colOriSelProduct_ As Collection
Dim colSearchResult_ As Collection
Dim colItemAdd_ As New Collection
Dim colItemDel_ As New Collection
Dim WorkDate_ As Date
Dim curCompanyID_ As String
Dim TruckTypeID_ As String
Dim callerFrm_ As Form
Public isJob As String
Public Status As String
Public ProjectID As String

Private Sub ChkAll_Click()
    cmdSearch_Click
End Sub

Private Sub cmdClose_Click()
GrOrderID.Clear
    Unload Me
End Sub
Public Property Get WorkDate() As Date
    WorkDate = WorkDate_
End Property
Public Property Let WorkDate(ByVal newval As Date)
    If Trim(WorkDate_) <> Trim(newval) Then
        WorkDate_ = Trim(newval)
    End If
End Property
Public Sub CmdSave_Click()
Dim x%
Dim tmpKey As String
    If colOriSelProduct_ Is Nothing Then Set colOriSelProduct_ = New Collection
    For x = 1 To colItemDel_.Count
        tmpKey = Trim("'" & colItemDel_(x).id & "_" & colItemDel_(x).QuotationID & "'")
        If isExist(colOriSelProduct_, tmpKey) Then
            callerFrm_.RemoveWasteData colItemDel_(x).id
        End If
    Next
    For x = 1 To colItemAdd_.Count
        tmpKey = Trim("'" & colItemAdd_(x).id & "'")
        If isExist(colOriSelProduct_, tmpKey) = False Then
            colOriSelProduct_.Add colItemAdd_(x), tmpKey
            callerFrm_.AddWasteDataItem colItemAdd_(x)
        End If
    Next
    GrOrderID.Clear
    Unload Me
End Sub

Private Sub cmdSearch_Click()
Dim tmpS As New BWGSearchManager
Dim tmpList As ListItem
Dim tmpWaste As BWGWasteDataCR
Dim tmpQuo As BWGQuotationHeader
Dim tmpArr
Dim criStr1 As String, tmpStr$, curNameSpec As String
Dim sqlStr$
Dim x%
Dim XX%
Dim i As Integer
Dim DetID As String
Dim DN, YN, MN, MN1, DD As String
    If Trim(txtKey.Text) = "" Then
        MsgBox "กรุณาใส่คำในการค้นหา", vbExclamation
        txtKey.SetFocus
        Exit Sub
    End If
    YN = Year(Date) + 543
    MN = Month(Date)
    If Len(MN) = 1 Then
        MN = "0" & MN
    End If
    DD = Format(Day(Date), "dd")
    DN = YN & "/" & MN & "/" & DD
    criStr1 = Replace(Trim(txtKey.Text), "*", "%")
    TxtD1.Text = Year(Date) + 540 & "/" & Month(Date) & "/" & Day(Date)
    
'    sqlStr = "(WasteNo Like '%" & criStr1 & "%' Or WasteName Like '%" & criStr1 & "%' Or QuotationNo Like '%" & criStr1 & "') And isTreatAble='Y' And QuotationNo<>'' And (isCancel<>'Y' Or isCancel is null) " 'and (LicenseExpDate>= '" & DN & "' or LicenseExpDate='' or LicenseExpDate is null) "
    If isEmpty(TruckTypeID_) Or IsNull(TruckTypeID_) Or TruckTypeID_ = "" Then
    sqlStr = "(WasteNo Like '%" & criStr1 & "%' Or WasteName Like '%" & criStr1 & "%' Or QuotationNo Like '%" & criStr1 & "') And QuotationNo<>'' And (isCancel<>'Y' Or isCancel is null)    " 'and (LicenseExpDate>= '" & DN & "' or LicenseExpDate='' or LicenseExpDate is null) "
    Else
    sqlStr = "(WasteNo Like '%" & criStr1 & "%' Or WasteName Like '%" & criStr1 & "%' Or QuotationNo Like '%" & criStr1 & "') And QuotationNo<>'' And (isCancel<>'Y' Or isCancel is null)  and truckTypeID='" & TruckTypeID_ & "'   " 'and (LicenseExpDate>= '" & DN & "' or LicenseExpDate='' or LicenseExpDate is null) "
    End If
    If Not curCustomer_ Is Nothing Then
        sqlStr = sqlStr & " And CustomerID='" & curCustomer_.id & "'"
         If curCustomer_.CustomerType <> "R" Then
            If Trim(curWasteGenNo_) <> "" Then
                sqlStr = sqlStr & " And wastemanageID='" & Trim(curWasteGenNo_) & "'"
            End If
            If Trim(curFactoryRegisNo_) <> "" Then
                sqlStr = sqlStr & " And FactoryRegisNo='" & Trim(curFactoryRegisNo_) & "'"
            End If
'            If Trim(curCompanyID_) <> "" Then
'                sqlStr = sqlStr & " And CompanyID='" & Trim(curCompanyID_) & "'"
'            End If
            If ChkAll.Value = 0 Then
                DetID = ""
                    If GrOrderID.Rows > 1 Then
                        If GrOrderID.TextMatrix(1, 0) <> "" Then
                                For i = 1 To GrOrderID.Rows - 1
                                    If DetID = "" Then
                                        If GrOrderID.Rows > 2 Then
                                            DetID = GrOrderID.TextMatrix(i, 0) & "'"
                                        Else
                                            DetID = GrOrderID.TextMatrix(i, 0)
                                        End If
                                    ElseIf i < GrOrderID.Rows - 1 Then
                                        DetID = DetID & ",'" & GrOrderID.TextMatrix(i, 0) & "'"
                                    Else
                                        DetID = DetID & ",'" & GrOrderID.TextMatrix(i, 0)
                                    End If
                                Next
                        End If
                    End If
                    If DetID <> "" Then
                        sqlStr = sqlStr & " AND quotationDetID in ('" & DetID & "') "
                    End If
            End If
            
            If Trim(isJob) <> "" Then
                sqlStr = sqlStr & " AND isJobs = '" & isJob & "'"
            End If
            If Trim(ProjectID) <> "" Then
            sqlStr = sqlStr & " and ProjectID='" & ProjectID & "' "
            End If
            If Status <> "" Then
                sqlStr = sqlStr & Status
            End If
        End If
       End If
    If TruckTypeID_ <> "" Then
        If Label3.Caption = "F042" Then
            Set colSearchResult_ = tmpS.JobAddWasteDataForTruckSearch(sqlStr)
        Else
            Set colSearchResult_ = tmpS.JobAddWasteDataForTruckSearch(sqlStr)
        End If
    Else
        If Label3.Caption = "F042" Then
            Set colSearchResult_ = tmpS.JobAddWasteDataSearch(sqlStr)
        Else
            Set colSearchResult_ = tmpS.JobAddWasteDataSearch(sqlStr)
        End If
    End If
 
        
        
    'Set colSearchResult_ = tmpS.WasteDataCRSearch(sqlStr)
    lvResult.ListItems.Clear
    For x = 1 To colSearchResult_.Count
        Set tmpWaste = colSearchResult_(x)
        'Set tmpList = lvResult.ListItems.Add(, "'" & tmpWaste.ID & "'", tmpWaste.WasteNo)
        Set tmpList = lvResult.ListItems.Add(, "'" & tmpWaste.id & "_" & tmpWaste.QuotationID & "'", tmpWaste.WasteNo)
        tmpList.SubItems(1) = tmpWaste.WasteName
        tmpList.SubItems(2) = tmpWaste.QuotationNo
        tmpList.SubItems(3) = tmpWaste.BMEQuotation
        tmpList.SubItems(4) = tmpWaste.TreatmentFee
        tmpList.SubItems(5) = tmpWaste.TreatmentUnitName
        If tmpWaste.IsDiscountPriceTrans Then
            tmpList.SubItems(6) = "Y"
        Else
            tmpList.SubItems(6) = "N"
        End If
       
        tmpList.SubItems(7) = tmpWaste.licenseStartDate
        tmpList.SubItems(8) = tmpWaste.LicenseExpDate
        If tmpWaste.isTrailer Then
            tmpList.SubItems(9) = tmpWaste.CarTypeName & " (พ่วง)"
        Else
            tmpList.SubItems(9) = tmpWaste.CarTypeName
        End If
        tmpList.SubItems(10) = tmpWaste.StaffCount
        tmpList.SubItems(11) = tmpWaste.StaffCharge
        If tmpWaste.isJob = "N" Then
            tmpList.SubItems(12) = "ใบเสนอราคาแบบ Routine"
        ElseIf tmpWaste.isJob = "Y" Then
            tmpList.SubItems(12) = "ใบเสนอราคาแบบ Jobs"
        ElseIf tmpWaste.isJob = "D" Then
            tmpList.SubItems(12) = "ใบเสนอราคาแบบ DOC"
        End If
        Select Case UCase(tmpWaste.WasteType)
            Case "HBL", "NBL", "BLF"
                tmpWaste.TSDFID = "F616DF9D156545BFBA7C256792F3136D"
                tmpList.SubItems(13) = "066200031"
            Case "HAZ", "NON", "BWG"
                tmpWaste.TSDFID = "C2C7E64E968F4EEAA537C877EE4127E3"
                tmpList.SubItems(13) = "056200025"
            Case "HIC", "HBI", "NIC", "NBI", "BBH", "BBL"
                tmpWaste.TSDFID = "F9B9701F2DF140EBA5B881649D0C9B66"
                tmpList.SubItems(13) = "085800027"
            Case Else
                tmpWaste.TSDFID = ""
                tmpList.SubItems(13) = ""
        End Select
          tmpList.SubItems(14) = tmpWaste.TransportFee
        tmpWaste.TSDFNo = tmpList.SubItems(13)
        tmpList.Checked = isExist(colOriSelProduct_, tmpList.Key) Or isExist(colItemAdd_, tmpList.Key)
         If tmpWaste.LicenseExpDate <> "" Then
            If tmpWaste.LicenseExpDate <= WorkDate_ Then
               tmpList.ForeColor = RGB(255, 0, 0)
               For XX = 1 To tmpList.ListSubItems.Count
                    tmpList.ListSubItems(XX).ForeColor = RGB(255, 0, 0)
              Next
              End If
        End If
        Set tmpList = Nothing
        Set tmpWaste = Nothing
    Next
    If lvResult.ListItems.Count = 0 Then
        MsgBox "ไม่พบข้อมูลที่ต้องการค้นหา", vbExclamation
    End If
    Label1(2).Caption = "ผลการค้นหา พบ " & lvResult.ListItems.Count & " รายการ"
End Sub

Private Function StatusDetail(Status As String) As String
        Select Case Status
            Case "P": StatusDetail = "ยังไม่สรุปผล"
            Case "A", "IA": StatusDetail = "รับดำเนินการ"
            Case "E": StatusDetail = "ไม่รับดำเนินการ"
            Case Else: StatusDetail = "ยังไม่สรุปผล"
        End Select
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then cmdSearch_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set colItemAdd_ = Nothing
    Set colItemDel_ = Nothing
End Sub

Private Sub lvResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'MsgBox ColumnHeader.Width
End Sub

Private Sub lvResult_ItemCheck(ByVal item As MSComctlLib.ListItem)
Dim tmpList As ListItem
Dim x%, selWasteID$
   Dim DN, YN, MN, MN1, DD As String
    YN = Year(Date) + 543
    MN = Month(Date)
    If Len(MN) = 1 Then
        MN = "0" & MN
    End If
    DD = Format(Day(Date), "dd")
   DN = YN & "/" & MN & "/" & DD

If item.Checked = True Then
    If item.SubItems(8) <> "" Then
      If item.SubItems(8) <= WorkDate_ Then
            MsgBox item.SubItems(1) & " หมดอายุแล้ว กรุณาตรวจสอบข้อมูล", vbOKOnly + vbCritical, "แจ้งเตือน สก2. หมดอายุ"
            item.Checked = False
            Exit Sub
        End If
       End If
    
        selWasteID = Trim(Left(item.Key, InStr(1, item.Key, "_") - 1))
        For x = 1 To lvSelected.ListItems.Count
            If Trim(selWasteID) = Trim(Left(lvSelected.ListItems(x).Key, InStr(1, lvSelected.ListItems(x).Key, "_") - 1)) Then
                MsgBox "ของเสียรายการนี้ถูกเลือกแล้วไม่สามารถเลือกซ้ำได้", vbExclamation
                item.Checked = False
                Exit Sub
            End If
        Next
        
        If lvSelected.ListItems.Count > 0 Then
            If lvSelected.ListItems(1).SubItems(4) <> item.SubItems(5) Then
             MsgBox "ของเสียรายการนี้มีหน่วยนับไม่ตรงกันไม่สามารถเลือกได้", vbExclamation
                item.Checked = False
                Exit Sub
            End If
        End If
        
        colItemAdd_.Add colSearchResult_(item.Key), item.Key
        If isExist(colItemDel_, item.Key) Then
            colItemDel_.Remove item.Key
        End If
            Set tmpList = lvSelected.ListItems.Add(, item.Key, item.Text)
            tmpList.SubItems(1) = item.SubItems(1)
            tmpList.SubItems(2) = item.SubItems(2)
            tmpList.SubItems(3) = item.SubItems(4)
            tmpList.SubItems(4) = item.SubItems(5)
            tmpList.SubItems(5) = item.SubItems(9)
            tmpList.SubItems(6) = item.SubItems(10)
            tmpList.SubItems(7) = item.SubItems(11)
            tmpList.SubItems(8) = item.SubItems(13)
            tmpList.SubItems(9) = item.SubItems(14)
           
           

    Else
        If isExist(colItemDel_, item.Key) = False Then
            colItemDel_.Add colSearchResult_(item.Key), item.Key
        End If
        If isExist(colItemAdd_, item.Key) Then
            colItemAdd_.Remove item.Key
        End If
        If isExist(lvSelected.ListItems, item.Key) = True Then
            lvSelected.ListItems.Remove item.Key
        End If
    End If
    Label1(1).Caption = "รายการสินค้าที่เลือก " & lvSelected.ListItems.Count & " รายการ"
End Sub

Public Sub SetCallerForm(frm As Form, selCust As BWGCustomer, curWasteGenNo As String, CompanyID As String, curFactoryRegisNo As String, TruckTypeID As String)
Dim x%
Dim tmpWaste As BWGWasteDataCR
Dim tmpList As ListItem
Dim tmpS As BWGSearchManager
Dim criStr As String
Dim strCond As String
    Set callerFrm_ = frm
    Set curCustomer_ = selCust
    lbCustomer.Caption = selCust.CustomerName
    Set colOriSelProduct_ = frm.getWastDataItems
    curCompanyID_ = Trim(CompanyID)
    TruckTypeID_ = TruckTypeID
    lvSelected.ListItems.Clear
    curWasteGenNo_ = Trim(curWasteGenNo)
    curFactoryRegisNo_ = Trim(curFactoryRegisNo)
    If Not colOriSelProduct_ Is Nothing Then
        For x = 1 To colOriSelProduct_.Count

'            Set colSearchResult_ = Nothing
'            Set colSearchResult_ = New Collection
'            Set tmpS = New BWGSearchManager
'            strCond = "(WasteDataID Like '%" & colOriSelProduct_(x).ID & "%')"
'            Set colSearchResult_ = tmpS.WasteDataCRSearch(strCond)
'            Set tmpS = Nothing
'            Set tmpWaste = colSearchResult_(1)
            Set tmpWaste = colOriSelProduct_(x)
            Set tmpList = lvSelected.ListItems.Add(, "'" & tmpWaste.id & "_" & tmpWaste.QuotationID & "'", tmpWaste.WasteNo)
            tmpList.SubItems(1) = tmpWaste.WasteName
            tmpList.SubItems(2) = tmpWaste.QuotationNo
            tmpList.SubItems(3) = tmpWaste.TreatmentFee
            tmpList.SubItems(4) = tmpWaste.TreatmentUnitName
            If tmpWaste.isTrailer Then
                tmpList.SubItems(5) = tmpWaste.CarTypeName & " (พ่วง)"
            Else
                tmpList.SubItems(5) = tmpWaste.CarTypeName
            End If
            tmpList.SubItems(6) = tmpWaste.StaffCount
            tmpList.SubItems(7) = tmpWaste.StaffCharge
            tmpList.SubItems(8) = tmpWaste.TSDFNo
            Set tmpList = Nothing
            Set tmpWaste = Nothing
        Next
    End If
    cmdSearch_Click
    Label1(1).Caption = "รายการที่เลือก " & lvSelected.ListItems.Count & " รายการ"
End Sub

Public Sub AddSelectedProduct(selWaste As BWGWasteDataCR)
Dim tmpKey As String
Dim tmpList As ListItem
    tmpKey = Trim("'" & selWaste.id & "'")
    If isExist(colItemAdd_, tmpKey) Then Exit Sub
    colItemAdd_.Add selWaste, tmpKey
    If isExist(colItemDel_, tmpKey) Then
        colItemDel_.Remove tmpKey
    End If
    If isExist(lvSelected.ListItems, tmpKey) = False Then
        Set tmpList = lvSelected.ListItems.Add(, tmpKey, selWaste.WasteNo)
        tmpList.SubItems(1) = selWaste.WasteName
        tmpList.SubItems(2) = selWaste.QuotationNo
        tmpList.SubItems(3) = selWaste.WasteType
        Set tmpList = Nothing
    End If
End Sub

Private Sub lvSelected_KeyDown(KeyCode As Integer, Shift As Integer)
Dim selResult As Integer
Dim tmpKey As String
    With lvSelected
        If KeyCode = 46 And Not .SelectedItem Is Nothing Then
            selResult = MsgBox("ต้องการลบกากของเสียนี้ออกจากใบเสนอราคาหรือไม่?", vbQuestion + vbYesNo)
            If selResult = vbYes Then
                tmpKey = Trim(.SelectedItem.Key)
                If isExist(colItemDel_, tmpKey) = False Then
                    If isExist(colOriSelProduct_, tmpKey) Then
                        colItemDel_.Add colOriSelProduct_(tmpKey), tmpKey
                    ElseIf isExist(colItemAdd_, tmpKey) Then
                        colItemDel_.Add colItemAdd_(tmpKey), tmpKey
                        colItemAdd_.Remove tmpKey
                    Else
                        MsgBox "ไม่พบรายการสินค้าที่ต้องการลบ", vbExclamation
                    End If
                End If
                If isExist(lvSelected.ListItems, tmpKey) = True Then
                    lvSelected.ListItems.Remove tmpKey
                End If
                If isExist(lvResult.ListItems, tmpKey) = True Then
                    lvResult.ListItems(tmpKey).Checked = False
                End If
            End If
        End If
    End With
End Sub

