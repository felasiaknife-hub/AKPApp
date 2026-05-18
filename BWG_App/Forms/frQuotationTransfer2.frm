VERSION 5.00
Begin VB.Form frmQuotationTransfer2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ย้ายใบเสนอราคา"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "ยืนยันการย้าย"
      Height          =   525
      Left            =   7860
      TabIndex        =   5
      Top             =   30
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ยกเลิก"
      Height          =   525
      Left            =   9630
      TabIndex        =   4
      Top             =   30
      Width           =   1695
   End
   Begin WasteManagment.ctlGrid grList 
      Height          =   6225
      Left            =   30
      TabIndex        =   3
      Top             =   600
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   10980
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   3630
      TabIndex        =   2
      Top             =   120
      Width           =   285
   End
   Begin VB.TextBox txtNewQuoNo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   5400
      Picture         =   "frQuotationTransfer2.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "เลขที่ใบเสนอราคาเดิม"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   1560
   End
End
Attribute VB_Name = "frmQuotationTransfer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim colJobDet_ As Collection
Dim selQuo_  As BWGQuotationHeader
Dim coltruckType_ As Collection
Dim curRow_ As Integer
Dim colNewQuo_ As Collection
Dim curQuoNo_ As String

Private Sub cmdConfirm_Click()
Dim x%, tmpJobDet As BWGJobDataDetail, Y%, iCount%
Dim tmpRec As New ADODB.Recordset, tmpCon As Connection, tmpQuoHead As BWGQuotationHeader
Dim tmpS As New BWGSearchManager, tmpCol As Collection, tmpQuoDet As BWGQuotationDetail, tmpColDet As Collection
    With grList.getGridObj
        For x = 1 To .Rows - 1
            If Trim(.TextMatrix(x, 0)) <> "" Then
                If Trim(.TextMatrix(x, 4)) = "" Then
                    MsgBox "กรุณาระบุ Action ในการดำเนินการให้ครบถ้วน", vbExclamation
                    Exit Sub
                ElseIf Trim(.TextMatrix(x, 4)) = "ย้าย" And Trim(.TextMatrix(x, 5)) = "" Then
                    MsgBox "กรุณาระบุใบเสนอราคาใหม่ที่ต้องการจะย้ายไปให้ครบถ้วน", vbExclamation
                    Exit Sub
                End If
            End If
        Next
    End With
    If MsgBox("ท่านต้องการยืนยันการย้ายรายการต่างๆ เหล่านี้หรือไม่?", vbQuestion + vbYesNo) = vbYes Then
        
        Set tmpColDet = tmpS.QuotationDetailSearch("CustomerID='" & selQuo_.CustomerID & "' And CompanyID='" & selQuo_.CompanyID & "'")
        Set tmpCol = tmpS.QuotationHeaderSearch("CustomerID='" & selQuo_.CustomerID & "' And CompanyID='" & selQuo_.CompanyID & "'")
        With grList.getGridObj
            For x = 1 To .Rows - 1
                If Trim(.TextMatrix(x, 0)) <> "" Then
                    Set tmpJobDet = colJobDet_("'" & Trim(.TextMatrix(x, 0)) & "'")
                    If Trim(.TextMatrix(x, 4)) = "ยกเลิก" Then
                        DBConnExc "Insert into tbJobDataDetail_Del (JobDetID,JobID,WasteDataID,EstWasteQty,WasteQtyUnitID,ContainerTypeDesc,ContainerCount,DetRemDesc,MarkForDelete,RowID,DestinationDesc,CarNo,TSDFID,QuotationID,ExtraLaborFee,TruckTypeID,TransportFee,TransferUnitID,TreatmentRate,TreatmentUnitID,LaborCount,isPriceIncTrans,MinWeightPerCar,NonMinRate,HazMinRate,NBLMinRate,HBLMinRate) (Select JobDetID,JobID,WasteDataID,EstWasteQty,WasteQtyUnitID,ContainerTypeDesc,ContainerCount,DetRemDesc,MarkForDelete,RowID,DestinationDesc,CarNo,TSDFID,QuotationID,ExtraLaborFee,TruckTypeID,TransportFee,TransferUnitID,TreatmentRate,TreatmentUnitID,LaborCount,isPriceIncTrans,MinWeightPerCar,NonMinRate,HazMinRate,NBLMinRate,HBLMinRate from tbJobDataDetail WHERE JobDetID='" & tmpJobDet.ID & "')"
                        DBConnExc "Update tbJobDataDetail_Del Set DelDate='" & TodayDate & "', DelStaffID='" & CurrentUser.ID & "' WHERE JobDetID='" & tmpJobDet.ID & "'"
                        DBConnExc "Delete tbJobDataDetail WHERE JobDetID='" & tmpJobDet.ID & "'"
                    ElseIf Trim(.TextMatrix(x, 4)) = "ย้าย" Then
                        GetDBConnection
                        For Y = 1 To tmpColDet.Count
                            Set tmpQuoDet = tmpColDet(Y)
                            If Trim(tmpQuoDet.QuotationID) = Trim(.TextMatrix(x, 7)) And Trim(tmpQuoDet.wastedataID) = (tmpJobDet.wastedataID) Then
                                Set tmpQuoHead = tmpCol("'" & tmpQuoDet.QuotationID & "'")
                                iCount = 0
                                DBConn.Execute "Update tbJobDataDetail_Del Set DelStaffID='" & CurrentUser.ID & "' WHERE JobDetID='" & tmpJobDet.ID & "'", iCount
                                If iCount = 0 Then
                                    DBConn.Execute "Insert into tbJobDataDetail_Del (JobDetID,JobID,WasteDataID,EstWasteQty,WasteQtyUnitID,ContainerTypeDesc,ContainerCount,DetRemDesc,MarkForDelete,RowID,DestinationDesc,CarNo,TSDFID,QuotationID,ExtraLaborFee,TruckTypeID,TransportFee,TransferUnitID,TreatmentRate,TreatmentUnitID,LaborCount,isPriceIncTrans,MinWeightPerCar,NonMinRate,HazMinRate,NBLMinRate,HBLMinRate) (Select JobDetID,JobID,WasteDataID,EstWasteQty,WasteQtyUnitID,ContainerTypeDesc,ContainerCount,DetRemDesc,MarkForDelete,RowID,DestinationDesc,CarNo,TSDFID,QuotationID,ExtraLaborFee,TruckTypeID,TransportFee,TransferUnitID,TreatmentRate,TreatmentUnitID,LaborCount,isPriceIncTrans,MinWeightPerCar,NonMinRate,HazMinRate,NBLMinRate,HBLMinRate from tbJobDataDetail WHERE JobDetID='" & tmpJobDet.ID & "')"
                                    DBConn.Execute "Update tbJobDataDetail_Del Set DelStaffID='" & CurrentUser.ID & "' WHERE JobDetID='" & tmpJobDet.ID & "'"
                                End If
                                DBConn.Execute "Update tbJobDataDetail Set QuotationID='" & tmpQuoHead.ID & "',ExtraLaborFee=" & tmpQuoHead.ExtraLaborFee & ",TruckTypeID='" & tmpQuoHead.TruckTypeID & "',TransportFee=" & tmpQuoHead.TransferRate & ",TransferUnitID='" & tmpQuoHead.TransferUnitID & "',TreatmentRate=" & tmpQuoDet.TreatmentRate & ",TreatmentUnitID='" & tmpQuoDet.TreatmentUnitID & "',LaborCount=" & tmpQuoHead.StaffCountPerCar & ",isPriceIncTrans='" & TF_To_YN(tmpQuoHead.IsPriceIncTransport) & "',MinWeightPerCar=" & tmpQuoHead.MinWeightPerCar & ",NonMinRate=" & tmpQuoHead.NonMinRate & ",HazMinRate=" & tmpQuoHead.HazMinRate & ",NBLMinRate=" & tmpQuoHead.NBLMinRate & ",HBLMinRate=" & tmpQuoHead.HBLMinRate & " WHERE JobDetID='" & tmpJobDet.ID & "'"
                                Set tmpQuoHead = Nothing
                                Set tmpQuoDet = Nothing
                                Exit For
                            End If
                            Set tmpQuoDet = Nothing
                        Next Y
                        CloseDBConnection
                    End If
                End If
            Next
        End With
        Set tmpCon = Nothing
        Set tmpCol = Nothing
        Set tmpColDet = Nothing
        MsgBox "ย้ายใบเสนอราคาเรียบร้อยแล้ว", vbInformation
        Unload Me
    End If
End Sub

Private Sub Command1_Click()
Dim newFrm As frmQuotation
    If Trim(txtNewQuoNo.Text) <> "" Then
        Set newFrm = New frmQuotation
        Load newFrm
        With newFrm
            .SetDataQuotation selQuo_
            .cmdPrint.Enabled = False
            .cmdApproved.Enabled = False
            .cmdAddItem.Enabled = False
            .Command1.Enabled = False
            .cmdSave.Enabled = False
            .cmdSearch.Enabled = False
            .cmdClear.Enabled = False
            .cmdCancel.Enabled = False
            .cmdCloseBill.Enabled = False
            .Show 1
        End With
        Set newFrm = Nothing
    Else
        MsgBox "ไม่มีเลขที่ใบเสนอราคา", vbExclamation
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    PopulateTruckType
    With grList.getGridObj
        .Cols = 8
        .Rows = 50
        .TextMatrix(0, 1) = "Job No"
        .TextMatrix(0, 2) = "กากของเสีย"
        .TextMatrix(0, 3) = "ประเภทรถ"
        .TextMatrix(0, 4) = "Action"
        .TextMatrix(0, 5) = "ใบเสนอราคาใหม่"
        .TextMatrix(0, 6) = ""
        .ColWidth(0) = 0
        .ColWidth(1) = 1290
        .ColWidth(2) = 3990
        .ColWidth(3) = 1425
        .ColWidth(4) = 855
        .ColWidth(5) = 3180
        .ColWidth(6) = 330
        .ColWidth(7) = 0
        .ColAlignment(1) = 4
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(6) = 4
    End With
    grList.ColInputType(4) = ComboBoxEnum
    grList.ColInputType(5) = ComboBoxEnum
End Sub

Private Sub grList_Click(Row As Integer, Col As Integer)
Dim newFrm As frmQuotation, tmpQuo As BWGQuotationDetail
Dim tmpCol As Collection, tmpS As New BWGSearchManager
    If Col = 6 And Trim(grList.getGridObj.TextMatrix(Row, 5)) <> "" Then
        With grList.getGridObj
            If Trim(.TextMatrix(Row, 5)) <> "" Then
                Set tmpCol = tmpS.QuotationHeaderSearch("QuotationID='" & Trim(.TextMatrix(Row, 7)) & "'")
                If tmpCol.Count > 0 Then
                    Set newFrm = New frmQuotation
                    Load newFrm
                    With newFrm
                        .SetDataQuotation tmpCol(1)
                        .cmdPrint.Enabled = False
                        .cmdApproved.Enabled = False
                        .cmdAddItem.Enabled = False
                        .Command1.Enabled = False
                        .cmdSave.Enabled = False
                        .cmdSearch.Enabled = False
                        .cmdClear.Enabled = False
                        .cmdCancel.Enabled = False
                        .cmdCloseBill.Enabled = False
                        .Show 1
                    End With
                    Set newFrm = Nothing
                    Set tmpCol = Nothing
                    Set tmpS = Nothing
                Else
                    MsgBox "ไม่มีข้อมูลใบเสนอราคา", vbExclamation
                    Exit Sub
                End If
            Else
                MsgBox "ไม่มีเลขที่ใบเสนอราคา", vbExclamation
            End If
        End With
    End If
End Sub

Public Sub SetDataQuotation(selQuo As BWGQuotationHeader)
    Set selQuo_ = selQuo
    txtNewQuoNo.Text = selQuo.QuotationNo
End Sub

Public Sub SetPendingJob(selCol As Collection, oldQuoNo As String)
Dim x%, tmpList As ListItem
Dim tmpJobDet As BWGJobDataDetail
    grList.ClearAllData
    Set colJobDet_ = selCol
    With grList.getGridObj
        .Rows = colJobDet_.Count + 20
        For x = 1 To selCol.Count
            Set tmpJobDet = selCol(x)
            .TextMatrix(x, 0) = Trim(tmpJobDet.ID)
            .TextMatrix(x, 1) = tmpJobDet.JobNo
            .TextMatrix(x, 2) = tmpJobDet.WasteNo & " : " & tmpJobDet.WasteName
            'If tmpJobDet.isTrailer Then
            '    .TextMatrix(x, 3) = tmpJobDet.QuoTruckTypeName & " (พ่วง)"
            'Else
            If isExist(coltruckType_, "'" & tmpJobDet.QuoTruckTypeID & "'") Then
                .TextMatrix(x, 3) = coltruckType_("'" & tmpJobDet.QuoTruckTypeID & "'").TypeDesc
            Else
                .TextMatrix(x, 3) = ""
            End If
            'End If
            Set tmpJobDet = Nothing
        Next
    End With
End Sub

Private Sub PopulateTruckType()
Dim tmpS As New BWGSearchManager
    Set coltruckType_ = tmpS.TruckTypeSearch
    Set tmpS = Nothing
End Sub

Private Sub grList_RolColChanged(Row As Integer, Col As Integer)
On Error GoTo ErrD
Dim tmpCbo As ComboBox, tmpS As New BWGSearchManager
Dim tmpQuo As BWGQuotationHeader, x%, tmpRec As New ADODB.Recordset
Dim tmpJobDet As BWGJobDataDetail
    If Col = 4 And Trim(grList.getGridObj.TextMatrix(Row, 0)) <> "" Then
        grList.ColEnabled(Col) = True
        Set tmpCbo = grList.getColObject(4)
        tmpCbo.Clear
        tmpCbo.AddItem "ย้าย"
        tmpCbo.AddItem "ยกเลิก"
        Set tmpCbo = Nothing
    ElseIf Col = 5 And Trim(grList.getGridObj.TextMatrix(Row, 0)) <> "" Then
        If Trim(grList.getGridObj.TextMatrix(Row, 4)) = "ยกเลิก" Then
            grList.ColEnabled(Col) = False
        Else
            grList.ColEnabled(Col) = True
            If curRow_ <> Row Or colNewQuo_ Is Nothing Then
                'ดึงใบเสนอราคาของลูกค้านั้นที่มี Waste ตัวนั้นและ Approve แล้วด้วย
                Set tmpJobDet = colJobDet_("'" & grList.getGridObj.TextMatrix(Row, 0) & "'")
                sqlStr = "SELECT * FROM vw_QuotationList WHERE CompanyID='" & selQuo_.CompanyID & "' And CustomerID='" & selQuo_.CustomerID & "' And WasteDataID='" & tmpJobDet.wastedataID & "' And QuotationNo<>'" & selQuo_.QuotationNo & "'"
                Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
                If tmpRec Is Nothing Then Exit Sub
                Set colNewQuo_ = New Collection
                Do Until tmpRec.EOF
                    Set tmpQuo = New BWGQuotationHeader
                    tmpQuo.QuotationNo = Trim("" & tmpRec("QuotationNo"))
                    tmpQuo.ID = Trim("" & tmpRec("QuotationID"))
                    tmpQuo.TruckTypeName = Trim("" & tmpRec("TruckTypeDesc"))
                    colNewQuo_.Add tmpQuo, "'" & tmpQuo.QuotationNo & "'"
                    Set tmpQuo = Nothing
                    tmpRec.MoveNext
                Loop
                Set tmpJobDet = Nothing
            End If
            Set tmpCbo = grList.getColObject(5)
            tmpCbo.Clear
            For x = 1 To colNewQuo_.Count
                tmpCbo.AddItem colNewQuo_(x).QuotationNo & " : " & colNewQuo_(x).TruckTypeName
            Next
        End If
    Else
        grList.ColEnabled(Col) = False
    End If
    curRow_ = Row
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Sub grList_UpdateValue(Row As Integer, Col As Integer)
Dim tmpVal As String
    tmpVal = Trim(grList.getGridObj.TextMatrix(Row, Col))
    If Col = 4 Then
        If Trim(tmpVal) = "ยกเลิก" Then
            grList.getGridObj.TextMatrix(Row, 5) = ""
            grList.getGridObj.TextMatrix(Row, 7) = ""
            grList.getGridObj.Col = 6
            Set grList.getGridObj.CellPicture = LoadPicture()
            grList.getGridObj.Text = ""
            grList.getGridObj.Col = 4
        End If
    ElseIf Col = 5 Then
        If Trim(tmpVal) <> "" Then
            grList.getGridObj.Col = 6
            grList.getGridObj.TextMatrix(Row, 7) = colNewQuo_(grList.getColObject(5).ListIndex + 1).ID
            grList.getGridObj.TextMatrix(Row, 4) = "ย้าย"
            Set grList.getGridObj.CellPicture = Me.Image1.Picture
            grList.getGridObj.Text = ""
            grList.getGridObj.Col = 5
        End If
    End If
End Sub
