VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmBoxTransferOut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Box Transfer Out"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10920
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6675
      Left            =   0
      TabIndex        =   2
      Top             =   -120
      Width           =   10995
      Begin VB.PictureBox PicGrid 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4275
         Left            =   4740
         ScaleHeight     =   4245
         ScaleWidth      =   6105
         TabIndex        =   16
         Top             =   840
         Width           =   6135
         Begin VB.CommandButton cmdSearch 
            Caption         =   "ค้นหา"
            Height          =   255
            Left            =   4500
            TabIndex        =   29
            Top             =   60
            Width           =   1095
         End
         Begin VB.TextBox txtSearch 
            Height          =   285
            Left            =   2160
            TabIndex        =   27
            Top             =   60
            Width           =   2235
         End
         Begin VSFlex7LCtl.VSFlexGrid VSGBoxList 
            Height          =   3435
            Left            =   0
            TabIndex        =   20
            Top             =   360
            Width           =   6075
            _cx             =   10716
            _cy             =   6059
            _ConvInfo       =   1
            Appearance      =   1
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
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
         Begin VB.PictureBox Picture2 
            BackColor       =   &H8000000A&
            Height          =   25
            Left            =   -60
            ScaleHeight     =   30
            ScaleWidth      =   6195
            TabIndex        =   19
            Top             =   3780
            Width           =   6195
         End
         Begin VB.PictureBox cmdPicGridCancel 
            BackColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   5400
            ScaleHeight     =   315
            ScaleWidth      =   615
            TabIndex        =   18
            Top             =   3840
            Width           =   675
         End
         Begin VB.PictureBox cmdAdd 
            BackColor       =   &H00C0FFC0&
            Height          =   375
            Left            =   4680
            ScaleHeight     =   315
            ScaleWidth      =   615
            TabIndex        =   17
            Top             =   3840
            Width           =   675
         End
         Begin VB.Label lblSearch 
            Caption         =   "ใส่เลขบ็อกที่ต้องการค้นหา"
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label lbl4 
            BackColor       =   &H00800080&
            Height          =   255
            Left            =   1320
            TabIndex        =   26
            ToolTipText     =   "จำหน่าย"
            Top             =   3900
            Width           =   255
         End
         Begin VB.Label lbl3 
            BackColor       =   &H000040C0&
            Height          =   255
            Left            =   1020
            TabIndex        =   25
            ToolTipText     =   "หาย"
            Top             =   3900
            Width           =   255
         End
         Begin VB.Label lbl2 
            BackColor       =   &H00FF8080&
            Height          =   255
            Left            =   720
            TabIndex        =   24
            ToolTipText     =   "ชำรุด"
            Top             =   3900
            Width           =   255
         End
         Begin VB.Label lbl1 
            BackColor       =   &H00008000&
            Height          =   255
            Left            =   420
            TabIndex        =   23
            ToolTipText     =   "พร้อมใช้งาน"
            Top             =   3900
            Width           =   255
         End
         Begin VB.Label lbl0 
            BackColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "จัดบ็อกให้ลูกค้าแล้ว"
            Top             =   3900
            Width           =   255
         End
      End
      Begin VB.CommandButton cmdSelectBox 
         Caption         =   "เลือก Box No."
         Height          =   435
         Left            =   7140
         TabIndex        =   15
         Top             =   1380
         Width           =   2115
      End
      Begin VB.ComboBox cboCarRegis 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtCustomer 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         Top             =   240
         Width           =   6435
      End
      Begin VB.Frame Frame2 
         Height          =   3195
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   10455
         Begin WasteManagment.ctlGrid ctlGrid1 
            Height          =   2955
            Left            =   60
            TabIndex        =   10
            Top             =   180
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   5212
         End
      End
      Begin VB.Frame Frame3 
         Height          =   675
         Left            =   240
         TabIndex        =   3
         Top             =   5880
         Width           =   10455
         Begin VB.CommandButton cmdCancel 
            Caption         =   "ยกเลิกการจัด"
            Height          =   375
            Left            =   6900
            TabIndex        =   33
            Top             =   180
            Width           =   1215
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "ออก"
            Height          =   375
            Left            =   9180
            TabIndex        =   32
            Top             =   180
            Width           =   1155
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "ลบข้อมูลการจัด"
            Enabled         =   0   'False
            Height          =   375
            Left            =   8340
            TabIndex        =   31
            Top             =   180
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "บันทึก"
            Height          =   375
            Left            =   5580
            TabIndex        =   30
            Top             =   180
            Width           =   1155
         End
         Begin VB.TextBox txtRemark 
            Height          =   375
            Left            =   960
            TabIndex        =   1
            Top             =   180
            Width           =   4215
         End
         Begin VB.Label Label1 
            Caption         =   "หมายเหตุ"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Label lblBoxType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "........"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3660
         TabIndex        =   21
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label lblTripControl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "........"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   300
         TabIndex        =   14
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "เลขที่ใบคุม"
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "ทะเบียน"
         Height          =   195
         Left            =   1860
         TabIndex        =   12
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "หมายเลขบ็อก"
         Height          =   195
         Left            =   7680
         TabIndex        =   9
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "ชื่อลูกค้า"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   -60
         X2              =   10920
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   8100
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label Label4 
         Caption         =   "ประเภทบ็อก"
         Height          =   195
         Left            =   3660
         TabIndex        =   5
         Top             =   1140
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmBoxTransferOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curCustomer_ As BWGCustomer
Dim myBox_ As BWGBoxTransactionH
Dim myJobOpen_ As BWGJobData
Dim JobDataCar As BWGJobDataTimeTable
Dim myJobOpenBox_ As BWGJobDataBox
Dim JobDataCarBox As BWGJobDataTimeTableBox
Dim colOriSelBox_ As Collection
Dim colSearchResult_ As Collection
Dim colItemAdd_ As New Collection
Dim colItemDel_ As New Collection
Dim BoxCount As Integer
Dim isDetailChange_ As Boolean
Dim OverAllCheck As Integer
Dim LastRowSel As Integer
Dim LastSearchSel As String
Dim TypeofTransaction As String

Public Sub TransactionType(TranssacType As String)
    TypeofTransaction = TranssacType
End Sub

Private Sub cboBoxType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("ยกเลิกทรานเซคชั่น ที่เกิดจากการจัดรถ (ที่มีการใส่จำนวน Box เข้าไปในระบบ)", vbOKCancel + vbExclamation) = vbCancel Then Exit Sub
    If TypeofTransaction = "JobData" Then
        If Not JobDataCar Is Nothing Then
            JobDataCar.UpdateBoxCanceled
'            lblStatus.Caption = "CANCELED"
            lblStatus.Visible = False
            cmdSave.Enabled = True
        End If
    ElseIf TypeofTransaction = "JobDataBox" Then
        If Not JobDataCarBox Is Nothing Then
            JobDataCarBox.UpdateBoxCanceled
'            lblStatus.Caption = "CANCELED"
            lblStatus.Visible = False
            cmdSave.Enabled = True
        End If
    End If
     If Not myBox_ Is Nothing Then
        If myBox_.TransactionStatus = "" Then Exit Sub
        myBox_.Delete 'ใช้การลบเพราะเปรียบเสมือนไม่มีทรานเซ็คชั่นเกิดขึ้น
    End If
    
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("ลบเฉพาะบ็อกที่จัดไปแล้ว ไม่ได้ลบใบคุมหรือใบเมนิเฟสแต่อย่างใด", vbOKCancel + vbExclamation) = vbCancel Then Exit Sub
    If Not myBox_ Is Nothing Then
        If myBox_.TransactionStatus = "" Then Exit Sub
        myBox_.Delete 'ใช้การลบเพราะเปรียบเสมือนไม่มีทรานเซ็คชั่นเกิดขึ้น
        lblStatus.Caption = "DELETED"
        lblStatus.Visible = True
        cmdSave.Enabled = False
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim x%, tmpKey As String
    If colOriSelBox_ Is Nothing Then Set colOriSelBox_ = New Collection
    For x = 1 To colItemDel_.Count
        tmpKey = Trim("'" & colItemDel_(x).BoxID & "'")
        If isExist(colOriSelBox_, tmpKey) Then
            RemoveBoxDataDetail colItemDel_(x).BoxID
        End If
    Next
    For x = 1 To colItemAdd_.Count
        tmpKey = Trim("'" & colItemAdd_(x).BoxID & "'")
        If isExist(colOriSelBox_, tmpKey) = False Then
            colOriSelBox_.Add colItemAdd_(x), tmpKey
            AddBoxDataDetail colItemAdd_(x)
        End If
    Next
    PicGrid.Visible = False
    myBox_.isChanged = True
    ShowBoxDetail
End Sub

Private Sub AddBoxDataDetail(ByVal selItem As BWGBoxDataItem)
    Dim tmpCol As Collection
    Dim x%, isNew As Boolean
    Dim tmpID As String
    Dim tmpBoxDetail As New BWGBoxTransactionD
    Set tmpCol = myBox_.TransactionDetails

    If tmpCol Is Nothing Then Set tmpCol = New Collection
    isNew = True
    tmpID = Trim(selItem.BoxID)
    For x = 1 To tmpCol.Count
        If tmpCol(x).BoxID = tmpID Then
            Set tmpBoxDetail = tmpCol(x)
            isNew = False
            Exit For
        End If
    Next
    
    If isNew Then
        If TypeofTransaction = "JobData" Then
            tmpBoxDetail.ID = myBox_.ID
            tmpBoxDetail.JobDataCarID = JobDataCar.JobDataCarID
            tmpBoxDetail.VendorTruckID = JobDataCar.VendorTruckID
            tmpBoxDetail.TripNo = JobDataCar.TripNo
            tmpBoxDetail.CarRegisID = JobDataCar.CarRegisNo
            tmpBoxDetail.TimeTableNo = JobDataCar.TimeTableNo
            tmpBoxDetail.DriverName = JobDataCar.DriverName
            tmpBoxDetail.BoxOwnersCode = selItem.BoxOwnersCode
            tmpBoxDetail.BoxType = selItem.BoxType
            tmpBoxDetail.BoxNo = selItem.BoxNo
            tmpBoxDetail.BoxID = selItem.BoxID
            tmpBoxDetail.isChanged = True
            tmpCol.Add tmpBoxDetail, "'" & selItem.BoxID & "'"
        ElseIf TypeofTransaction = "JobDataBox" Then
            tmpBoxDetail.ID = myBox_.ID
            tmpBoxDetail.JobDataCarID = JobDataCarBox.JobDataCarID
            tmpBoxDetail.VendorTruckID = JobDataCarBox.VendorTruckID
            tmpBoxDetail.TripNo = JobDataCarBox.TripNo
            tmpBoxDetail.CarRegisID = JobDataCarBox.CarRegisNo
            tmpBoxDetail.TimeTableNo = JobDataCarBox.TimeTableNo
            tmpBoxDetail.DriverName = JobDataCarBox.DriverName
            tmpBoxDetail.BoxOwnersCode = selItem.BoxOwnersCode
            tmpBoxDetail.BoxType = selItem.BoxType
            tmpBoxDetail.BoxNo = selItem.BoxNo
            tmpBoxDetail.BoxID = selItem.BoxID
            tmpBoxDetail.isChanged = True
            tmpCol.Add tmpBoxDetail, "'" & selItem.BoxID & "'"
        End If
    End If

    
    isDetailChange_ = True
    myBox_.isChanged = True
    Set tmpBoxDetail = Nothing
    
    If myBox_.TransactionDetails Is Nothing Then
        Set myBox_.TransactionDetails = tmpCol
        isDetailChange_ = True
        myBox_.isChanged = True
    End If
    Set tmpCol = Nothing
End Sub

Public Sub RemoveBoxDataDetail(selID As String)
Dim tmpCol As Collection
Dim tmpKey As String, x%
    Set tmpCol = myBox_.TransactionDetails
    If tmpCol Is Nothing Then Exit Sub
    For x = 1 To tmpCol.Count
        If tmpCol(x).BoxID = selID Then
            tmpKey = Trim("'" & tmpCol(x).BoxID & "'")
            Exit For
        End If
    Next
    If isExist(tmpCol, tmpKey) Then
        tmpCol.Remove tmpKey
        isDetailChange_ = True
        myBox_.isChanged = True
    End If
End Sub

Private Function GetLastRow()
    Dim i As Integer
    For i = 1 To ctlGrid1.getGridObj.Rows
        If Trim(ctlGrid1.getGridObj.TextMatrix(i, 0)) = "" Then
            GetLastRow = i
            Exit For
        End If
    Next
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPicGridCancel_Click()
    PicGrid.Visible = False
    OverAllCheck = 0
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrHanler
    If myBox_.TransactionDetails.Count <= 0 Then MsgBox "โปรดเลือกบ็อกอย่างน้อย 1 รายการ", vbExclamation: Exit Sub
    If myBox_.isChanged = False Then Exit Sub
    If TypeofTransaction = "JobData" Then
        With myBox_
            .JobDataCarID = JobDataCar.JobDataCarID
            .TimeTableNo = JobDataCar.TimeTableNo
            .VendorTruckID = JobDataCar.VendorTruckID
            If OverAllCheck = JobDataCar.BoxCount Then
                .TransactionStatus = "COMPLETE"
            ElseIf OverAllCheck < JobDataCar.BoxCount Then
                .TransactionStatus = "INCOMPLETE"
            End If
            .TransactionType = "TFO"
            .TransactionDesc = Trim(txtRemark.Text)
            .Save
        End With
    ElseIf TypeofTransaction = "JobDataBox" Then
        With myBox_
            .JobDataCarID = JobDataCarBox.JobDataCarID
            .TimeTableNo = JobDataCarBox.TimeTableNo
            .VendorTruckID = JobDataCarBox.VendorTruckID
            If OverAllCheck = JobDataCarBox.BoxCount Then
                .TransactionStatus = "COMPLETE"
            ElseIf OverAllCheck < JobDataCarBox.BoxCount Then
                .TransactionStatus = "INCOMPLETE"
            End If
            .TransactionType = "TFO"
            .TransactionDesc = Trim(txtRemark.Text)
            .Save
        End With
    End If
    MsgBox "บันทึกสำเร็จ!!", vbInformation
    Exit Sub
    
ErrHanler:
    MsgBox "เกิดข้อผิดพลาด กรุณาติดต่อผู้พัฒนาระบบ", vbInformation
    Exit Sub
End Sub

Private Sub cmdSearch_Click()
    Dim Row As Integer
    If LastRowSel = 0 Then VSGBoxList.Select 0, 0, 0, 0
    For Row = LastRowSel To VSGBoxList.Rows - 1
        If InStr(1, VSGBoxList.TextMatrix(Row, 6), txtSearch) > 0 Then
            VSGBoxList.Select Row, 0, Row, VSGBoxList.Cols - 1
            LastRowSel = Row + 1
            VSGBoxList.SelectionMode = flexSelectionByRow
            VSGBoxList.ScrollTrack = False
            VSGBoxList.ScrollTips = True
            VSGBoxList.TopRow = Row
            VSGBoxList.Refresh
            Exit For
        End If
    Next Row
    If Row = VSGBoxList.Rows Then
        MsgBox "Next find not found.", vbInformation
        LastRowSel = 0
    End If
End Sub

Private Sub cmdSelectBox_Click()
    isDetailChange_ = False
    OverAllCheck = 0
    Set colOriSelBox_ = GetBoxDataItem
    PicGrid.Visible = True
    If isDetailChange_ Then
        ShowBoxDetail
    End If
End Sub

Public Function GetBoxDataItem() As Collection
    Dim x%
    Dim tmpCol As New Collection
    Dim tmpDet As BWGBoxTransactionD
    Dim tmpBoxItem As BWGBoxDataItem
    Dim selCol As Collection, tmpStr As String, tmpArr
    If myBox_ Is Nothing Then Exit Function
    myBox_.TransactionType = "TFO"
    Set selCol = myBox_.TransactionDetails
    With selCol
        For x = 1 To .Count
            Set tmpDet = .item(x)
            If Trim(tmpDet.BoxNo) <> "" Then
                Set tmpBoxItem = New BWGBoxDataItem
                tmpBoxItem.BoxOwnersCode = tmpDet.BoxOwnersCode
                tmpBoxItem.BoxType = tmpDet.BoxType
                tmpBoxItem.BoxNo = tmpDet.BoxNo
                tmpBoxItem.BoxID = tmpDet.BoxID
                tmpCol.Add tmpBoxItem, "'" & tmpDet.BoxOwnersCode & tmpDet.BoxType & tmpDet.BoxNo & "'"
                OverAllCheck = OverAllCheck + 1
                Set tmpBoxItem = Nothing
            End If
            Set tmpDet = Nothing
        Next
    End With
    
    Set GetBoxDataItem = tmpCol
    Set tmpCol = Nothing
End Function

Private Sub Form_Resize()
    On Error Resume Next
    Frame1.Width = Me.Width
    Frame2.Width = Me.Width - (Frame2.Left * 2.6)
    Frame3.Width = Me.Width - (Frame3.Left * 2.6)
    
    lblStatus.Left = Me.Width - lblStatus.Width - 200
    Line1.X2 = Me.Width
    
    ctlGrid1.Width = Frame2.Width
    cmdSave.Left = Frame3.Width - (cmdSave.Width * 4.8)
    cmdCancel.Left = Frame3.Width - (cmdCancel.Width * 3.6)
    cmdDelete.Left = Frame3.Width - (cmdDelete.Width * 2.4)
    cmdExit.Left = Frame3.Width - (cmdExit.Width * 1.2)
    txtRemark.Width = Frame3.Width - (cmdExit.Width * 6)
End Sub

Private Sub Form_Load()
    OverAllCheck = 0
    PicGrid.Top = cmdSelectBox.Top
    PicGrid.Visible = False
    SetGrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set colItemAdd_ = Nothing
    Set colItemDel_ = Nothing
    LastRowSel = 0
    LastSearchSel = ""
    Unload Me
End Sub

Private Sub SetGrid()
    With ctlGrid1.getGridObj
        .Cols = 7
        .Rows = 6
        .TextMatrix(0, 0) = "JobDataCarID"
        .TextMatrix(0, 1) = "TripNo"
        .TextMatrix(0, 2) = "ใบคุม"
        .TextMatrix(0, 3) = "ทะเบียน"
        .TextMatrix(0, 4) = "Box Owner"
        .TextMatrix(0, 5) = "Box Type"
        .TextMatrix(0, 6) = "Box No."
        .ColWidth(0) = 0
        .ColWidth(1) = 0
        .ColWidth(2) = 1500
        .ColWidth(3) = 2000
        .ColWidth(4) = 2000
        .ColWidth(5) = 2000
        .ColWidth(6) = 1500
        
        .ColAlignment(0) = 4
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 4
        .ColAlignment(4) = 4
        .ColAlignment(5) = 4
        .ColAlignment(6) = 4
    End With
    ctlGrid1.ColEnabled(0) = False
    ctlGrid1.ColEnabled(1) = False
    ctlGrid1.ColEnabled(2) = False
    ctlGrid1.ColEnabled(3) = False
    ctlGrid1.ColEnabled(4) = False
    ctlGrid1.ColEnabled(5) = False
    ctlGrid1.ColEnabled(6) = False
End Sub

Public Sub SetDataBoxOpen(ByVal selJob As BWGJobData, ByVal selCar As BWGJobDataTimeTable)
Dim tmpStr
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    Set myJobOpen_ = selJob
    With myJobOpen_
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1), False
        Set JobDataCar = selCar
        SearchBoxType JobDataCar.TimeTableNo, JobDataCar.BoxType, JobDataCar.CarRegisNo
        InitialBoxData JobDataCar.JobDataCarID, JobDataCar.TimeTableNo, JobDataCar.VendorTruckID, JobDataCar.BoxCount
    End With
End Sub

Public Sub SetDataBoxOpenBox(ByVal selJob As BWGJobDataBox, ByVal selCar As BWGJobDataTimeTableBox)
Dim tmpStr
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    Set myJobOpenBox_ = selJob
    With myJobOpenBox_
        Set tmpCol = tmpS.CustomerSearch("CustomerID='" & .CustomerID & "'")
        If tmpCol.Count > 0 Then SetDataCustomer tmpCol(1), False
        Set JobDataCarBox = selCar
        SearchBoxType JobDataCarBox.TimeTableNo, JobDataCarBox.BoxType, JobDataCarBox.CarRegisNo
        InitialBoxData JobDataCarBox.JobDataCarID, JobDataCarBox.TimeTableNo, JobDataCarBox.VendorTruckID, JobDataCarBox.BoxCount
    End With
End Sub

Private Sub SearchBoxType(ByVal TimeTableNo As String, ByVal BoxType As String, ByVal CarRegisNo As String)
    Dim RS As ADODB.Recordset
    Dim BoxItem As New BWGBoxDataItem
    Dim x As Integer, Col As Long, tSql As String
    lblBoxType.Caption = ""
    lblTripControl.Caption = TimeTableNo
    cboCarRegis.Clear
    cboCarRegis.AddItem CarRegisNo
    
    tSql = "SELECT BoxType, BoxTypeDesc FROM Box_BoxType WHERE BoxType = '" & BoxType & "'"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockOptimistic, DBConStock)
    If Not RS.EOF Then
        lblBoxType.Caption = RS!BoxTypeDesc
    End If
    Set RS = Nothing
    
    If cboCarRegis.ListCount > 0 Then cboCarRegis.ListIndex = 0
    With VSGBoxList
        .Clear flexClearData, flexClearText
        .Cols = 7
        .Rows = 2
        .FixedCols = 0
        .TextMatrix(0, 0) = "•"
        .TextMatrix(0, 1) = "Owner"
        .TextMatrix(0, 2) = "Box Type"
        .TextMatrix(0, 3) = "Box No."
        .TextMatrix(0, 4) = "Status"
        .TextMatrix(0, 5) = "Desc"
        .TextMatrix(0, 6) = "BoxID"
        .ColAlignment(0) = flexAlignCenterTop
        .ColDataType(0) = flexDTBoolean
        .ColDataType(1) = flexDTString
        .ColDataType(2) = flexDTString
        .ColDataType(3) = flexDTString
        .ColDataType(4) = flexDTString
        .ColDataType(5) = flexDTString
        .ColDataType(6) = flexDTString
        
        .ColWidth(0) = 800
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 0
        
        .Editable = flexEDKbdMouse
    End With
    Dim BoxStatus As String
    Dim BColor
    tSql = "BoxType = '" & BoxType & "'"
    Set colSearchResult_ = BoxItem.BoxDataItemSearch(tSql)
    For x = 1 To colSearchResult_.Count
        Set BoxItem = colSearchResult_(x)
        With VSGBoxList
            .Rows = colSearchResult_.Count + 1
            .TextMatrix(x, 0) = "0"
            .TextMatrix(x, 1) = BoxItem.BoxOwnersCode
            .TextMatrix(x, 2) = BoxItem.BoxType
            .TextMatrix(x, 3) = BoxItem.BoxNo
            Select Case BoxItem.BoxStatus
                Case 0 'บ็อกนี้ถูกจัดให้ลูกค้าแล้ว
                    BoxStatus = "บ็อกนี้ถูกจัดให้ลูกค้าแล้ว"
                    BColor = lbl0.BackColor
                Case 1 'พร้อมใช้
                    BoxStatus = "พร้อมใช้"
                    BColor = lbl1.BackColor
                Case 2 'ชำรุด
                    BoxStatus = "ชำรุด"
                    BColor = lbl2.BackColor
                Case 3 'หาย
                    BoxStatus = "หาย"
                    BColor = lbl3.BackColor
                Case 4 'จำหน่าย
                    BoxStatus = "จำหน่าย"
                    BColor = lbl4.BackColor
            End Select
            .TextMatrix(x, 4) = BoxStatus
            .TextMatrix(x, 5) = BoxItem.BoxDesc
            .TextMatrix(x, 6) = BoxItem.BoxID
            .Cell(flexcpForeColor, x, 0, x, .Cols - 1) = BColor
        End With
    Next x
    Set BoxItem = Nothing
End Sub

Private Sub InitialBoxData(ByVal JobDataCarID As String, ByVal TimeTableNo As String, ByVal VendorTruckID As String, ByVal BoxCount As Integer)
'On Error GoTo ErrD
    Dim RS As ADODB.Recordset
    Dim tSql As String, i As Integer
    Dim tmpS As New BWGBoxTransactionH
    Dim tmpCol As Collection
    Set myBox_ = New BWGBoxTransactionH
    Set tmpCol = tmpS.TransactionHeaderSearch("JobDataCarID = '" & JobDataCarID & "' AND TimeTableNo = '" & TimeTableNo & "' AND VendorTruckID = '" & VendorTruckID & "'", "TFO")
    Set colItemAdd_ = Nothing
    Set colItemDel_ = Nothing
    If tmpCol.Count > 0 Then
        Set myBox_ = tmpCol(1)
        lblStatus = Trim(myBox_.TransactionStatus)
        lblStatus.Visible = True
        txtRemark.Text = myBox_.TransactionDesc
        ShowBoxDetail
    Else
        CreateNewBoxOpen JobDataCarID, BoxCount
    End If
    Set tmpCol = Nothing
    Exit Sub
ErrD:
    MsgBox err.Description
    Exit Sub
End Sub

Private Sub ShowBoxDetail()
    Dim tmpBoxDetail As BWGBoxTransactionD
    Dim Index As Integer, tmpList As ListItem
    BoxCount = 0
    ctlGrid1.getGridObj.Clear
    SetGrid
    For Index = 1 To myBox_.TransactionDetails.Count
        Set tmpBoxDetail = myBox_.TransactionDetails(Index)
        With ctlGrid1.getGridObj
            .Rows = .Rows + Index
            .TextMatrix(Index, 0) = tmpBoxDetail.JobDataCarID
            .TextMatrix(Index, 1) = tmpBoxDetail.TripNo
            .TextMatrix(Index, 2) = tmpBoxDetail.TimeTableNo
            .TextMatrix(Index, 3) = tmpBoxDetail.CarRegisID
            .TextMatrix(Index, 4) = tmpBoxDetail.BoxOwnersCode
            .TextMatrix(Index, 5) = tmpBoxDetail.BoxType
            .TextMatrix(Index, 6) = tmpBoxDetail.BoxNo
        End With
        Set tmpBoxDetail = Nothing
    Next
End Sub

Private Sub CreateNewBoxOpen(ByVal JobDataCarID As String, ByVal BoxCount As Integer)
    Set myBox_ = Nothing
    Set myBox_ = New BWGBoxTransactionH
    With myBox_
        .ID = GetGUID
        .JobDataCarID = JobDataCarID
        .TransactionDate = TodayDate
        .TransactionStatus = "TFO"
        .TransactionStatusDate = TodayDate
        .TransactionDesc = "TFO"
        .isChanged = False
    End With
    lblStatus = "จำนวน Box ที่ต้องจัด " & BoxCount
    lblStatus.Visible = True
    txtRemark.Text = ""
End Sub

Public Sub SetTextToCombo(selTxt As String, selCbo As ComboBox)
On Error GoTo setErr:
    selCbo.Text = Trim(selTxt)
    Exit Sub
setErr:
    selCbo.ListIndex = -1
    Exit Sub
End Sub

Private Function isExist(tmpCol As Object, strKey As String, Optional LView As Boolean = False) As Boolean
On Error GoTo genErr:
Dim tmpObj As Object
Dim i As Long
    If LView = False Then
        'If tmpCol(strKey).BoxID <> "" Then isExist = True
        Set tmpObj = tmpCol(strKey)
        isExist = True
        Exit Function
    Else
        For i = 1 To tmpCol.ListItems.Count
            If tmpCol.ListItems(i).Key = strKey Then isExist = True: Exit For
        Next
    End If
    Exit Function
genErr:
    isExist = False
End Function

Public Sub SetDataCustomer(selCustomer As BWGCustomer, Optional isAssigned As Boolean = True)
    On Error GoTo ErrD
    Dim tmpRec As New ADODB.Recordset
    Dim x%, sqlStr As String
    Set curCustomer_ = Nothing
    Set curCustomer_ = selCustomer
    If curCustomer_ Is Nothing Then Exit Sub
    txtCustomer.Text = curCustomer_.CustomerName
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub






Private Sub txtRemark_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys vbTab
End Sub

Private Sub txtRemark_Validate(Cancel As Boolean)
    If Not myBox_ Is Nothing Then
        myBox_.TransactionDesc = Trim(txtRemark.Text)
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And Trim(txtSearch) <> "" Then
        If LastSearchSel <> Trim(txtSearch) Then LastRowSel = 0
        cmdSearch_Click
        txtSearch.SelStart = 0
        txtSearch.SelLength = Len(txtSearch)
        LastSearchSel = Trim(txtSearch)
    End If
End Sub

Private Sub VSGBoxList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo ErrHanler
    Dim x%, selBoxNo As String
    If Col > 0 Then Exit Sub
    If VSGBoxList.TextMatrix(Row, 6) = "" Then Exit Sub
    If CBool(VSGBoxList.TextMatrix(Row, Col)) = True Then
        selBoxNo = "'" & VSGBoxList.TextMatrix(Row, 6) & "'"
        For x = 1 To VSGBoxList.Rows - 1
            If isExist(colItemAdd_, selBoxNo) = True Then
                MsgBox "Box ใบนี้ถูกเลือกแล้วไม่สามารถเลือกซ้ำได้", vbExclamation
                VSGBoxList.TextMatrix(Row, Col) = "0"
                selBoxNo = ""
                Exit Sub
            End If
        Next
        
        OverAllCheck = OverAllCheck + 1
        If TypeofTransaction = "JobData" Then
            If OverAllCheck > JobDataCar.BoxCount Then
                OverAllCheck = OverAllCheck - 1
                MsgBox "ไม่สามารถเลือกบ็อกได้มากกว่าจำนวนที่กำหนดไว้แล้ว", vbInformation
                VSGBoxList.TextMatrix(Row, Col) = "0"
                Exit Sub
            End If
        ElseIf TypeofTransaction = "JobDataBox" Then
            If OverAllCheck > JobDataCarBox.BoxCount Then
                OverAllCheck = OverAllCheck - 1
                MsgBox "ไม่สามารถเลือกบ็อกได้มากกว่าจำนวนที่กำหนดไว้แล้ว", vbInformation
                VSGBoxList.TextMatrix(Row, Col) = "0"
                Exit Sub
            End If
        End If
        
        If VSGBoxList.TextMatrix(Row, 4) <> "พร้อมใช้" Then
            OverAllCheck = OverAllCheck - 1
            MsgBox "บ็อกไม่พร้อมใช้งาน", vbInformation
            VSGBoxList.TextMatrix(Row, Col) = "0"
            Exit Sub
        End If
                
        colItemAdd_.Add colSearchResult_(selBoxNo), selBoxNo
        If isExist(colItemDel_, selBoxNo) Then
            colItemDel_.Remove selBoxNo
        End If
    Else
        selBoxNo = "'" & VSGBoxList.TextMatrix(Row, 6) & "'"
        VSGBoxList.TextMatrix(Row, Col) = "0"
        OverAllCheck = OverAllCheck - 1
        If OverAllCheck = 0 Then OverAllCheck = 0
    
        If isExist(colItemDel_, selBoxNo) = False Then
            colItemDel_.Add colSearchResult_(selBoxNo), selBoxNo
        End If
        If isExist(colItemAdd_, selBoxNo) Then
            colItemAdd_.Remove selBoxNo
        End If
    End If
Exit Sub
ErrHanler:
    Exit Sub
End Sub

Private Sub VSGBoxList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col > 0 Then Cancel = True: Exit Sub
End Sub

Private Sub VSGBoxList_BeforeScrollTip(ByVal Row As Long)
    VSGBoxList.ScrollTipText = " Row " & Row & ": " & VSGBoxList.Cell(flexcpTextDisplay, Row, 0) & " "
End Sub
