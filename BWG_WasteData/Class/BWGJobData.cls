VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGJobData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim jobNo_ As String
Dim customerID_ As String
Dim customerName_ As String
Dim customerSiteID_ As String
Dim customerSiteName_ As String
Dim customerSiteTel_ As String
Dim customerSiteFax_ As String
Dim customerSiteStaff_ As String
Dim customerSiteAumphurID_ As String
Dim customerSiteProvinceID_ As String
Dim customerSiteProvinceName_ As String
Dim workDate_ As String
Dim workTime_ As String
Dim issuedStaffName_ As String
Dim issuedStaffID_ As String
Dim saleStaffName_ As String
Dim remDesc_ As String
Dim issuedDate_ As String
Dim companyID_ As String
Dim isCancel_ As String
Dim wasteCreateNo_ As String
Dim wasteType_ As String
Dim colJobDet_ As Collection
Dim colJobCar_ As Collection
Dim menifestCountAll_ As Integer
Dim menifestCountIssued_ As Integer
Dim underStaffID_ As String
Dim underStaffName_ As String
Dim isCompleted_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get JobNo() As String
    JobNo = jobNo_
End Property
Public Property Let JobNo(ByVal newVal As String)
    If Trim(jobNo_) <> Trim(newVal) Then
        jobNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerID() As String
    CustomerID = customerID_
End Property
Public Property Let CustomerID(ByVal newVal As String)
    If Trim(customerID_) <> Trim(newVal) Then
        customerID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerName() As String
    CustomerName = customerName_
End Property
Public Property Let CustomerName(ByVal newVal As String)
    If Trim(customerName_) <> Trim(newVal) Then
        customerName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteID() As String
    CustomerSiteID = customerSiteID_
End Property
Public Property Let CustomerSiteID(ByVal newVal As String)
    If Trim(customerSiteID_) <> Trim(newVal) Then
        customerSiteID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteName() As String
    CustomerSiteName = customerSiteName_
End Property
Public Property Let CustomerSiteName(ByVal newVal As String)
    If Trim(customerSiteName_) <> Trim(newVal) Then
        customerSiteName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteTel() As String
    CustomerSiteTel = customerSiteTel_
End Property
Public Property Let CustomerSiteTel(ByVal newVal As String)
    If Trim(customerSiteTel_) <> Trim(newVal) Then
        customerSiteTel_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteFax() As String
    CustomerSiteFax = customerSiteFax_
End Property
Public Property Let CustomerSiteFax(ByVal newVal As String)
    If Trim(customerSiteFax_) <> Trim(newVal) Then
        customerSiteFax_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteStaff() As String
    CustomerSiteStaff = customerSiteStaff_
End Property
Public Property Let CustomerSiteStaff(ByVal newVal As String)
    If Trim(customerSiteStaff_) <> Trim(newVal) Then
        customerSiteStaff_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteAumphurID() As String
    CustomerSiteAumphurID = customerSiteAumphurID_
End Property
Public Property Let CustomerSiteAumphurID(ByVal newVal As String)
    If Trim(customerSiteAumphurID_) <> Trim(newVal) Then
        customerSiteAumphurID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteProvinceID() As String
    CustomerSiteProvinceID = customerSiteProvinceID_
End Property
Public Property Let CustomerSiteProvinceID(ByVal newVal As String)
    If Trim(customerSiteProvinceID_) <> Trim(newVal) Then
        customerSiteProvinceID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteProvinceName() As String
    CustomerSiteProvinceName = customerSiteProvinceName_
End Property
Public Property Let CustomerSiteProvinceName(ByVal newVal As String)
    If Trim(customerSiteProvinceName_) <> Trim(newVal) Then
        customerSiteProvinceName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WorkDate() As String
    WorkDate = workDate_
End Property
Public Property Let WorkDate(ByVal newVal As String)
    If Trim(workDate_) <> Trim(newVal) Then
        workDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WorkTime() As String
    WorkTime = workTime_
End Property
Public Property Let WorkTime(ByVal newVal As String)
    If Trim(workTime_) <> Trim(newVal) Then
        workTime_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedStaffName() As String
    IssuedStaffName = issuedStaffName_
End Property
Public Property Let IssuedStaffName(ByVal newVal As String)
    If Trim(issuedStaffName_) <> Trim(newVal) Then
        issuedStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedStaffID() As String
    IssuedStaffID = issuedStaffID_
End Property
Public Property Let IssuedStaffID(ByVal newVal As String)
    If Trim(issuedStaffID_) <> Trim(newVal) Then
        issuedStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SaleStaffName() As String
    SaleStaffName = saleStaffName_
End Property
Public Property Let SaleStaffName(ByVal newVal As String)
    If Trim(saleStaffName_) <> Trim(newVal) Then
        saleStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get RemDesc() As String
    RemDesc = remDesc_
End Property
Public Property Let RemDesc(ByVal newVal As String)
    If Trim(remDesc_) <> Trim(newVal) Then
        remDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedDate() As String
    IssuedDate = issuedDate_
End Property
Public Property Let IssuedDate(ByVal newVal As String)
    If Trim(issuedDate_) <> Trim(newVal) Then
        issuedDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get companyID() As String
    companyID = companyID_
End Property
Public Property Let companyID(ByVal newVal As String)
    If Trim(companyID_) <> Trim(newVal) Then
        companyID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isCancel() As Boolean
    If Trim(isCancel_) = "Y" Then
        isCancel = True
    Else
        isCancel = False
    End If
End Property
Public Property Let isCancel(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isCancel_) <> Trim(tmpStr) Then
        isCancel_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get WasteCreateNo() As String
    WasteCreateNo = wasteCreateNo_
End Property
Public Property Let WasteCreateNo(ByVal newVal As String)
    If Trim(wasteCreateNo_) <> Trim(newVal) Then
        wasteCreateNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteType() As String
    WasteType = wasteType_
End Property
Public Property Let WasteType(ByVal newVal As String)
    If Trim(wasteType_) <> Trim(newVal) Then
        wasteType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestCountAll() As Integer
    MenifestCountAll = menifestCountAll_
End Property
Public Property Let MenifestCountAll(ByVal newVal As Integer)
    If menifestCountAll_ <> newVal Then
        menifestCountAll_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get MenifestCountIssued() As Integer
    MenifestCountIssued = menifestCountIssued_
End Property
Public Property Let MenifestCountIssued(ByVal newVal As Integer)
    If menifestCountIssued_ <> newVal Then
        menifestCountIssued_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get UnderStaffID() As String
    UnderStaffID = underStaffID_
End Property
Public Property Let UnderStaffID(ByVal newVal As String)
    If Trim(underStaffID_) <> newVal Then
        underStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get UnderStaffName() As String
    UnderStaffName = underStaffName_
End Property
Public Property Let UnderStaffName(ByVal newVal As String)
    If Trim(underStaffName_) <> newVal Then
        underStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isCompleted() As Boolean
    If Trim(isCompleted_) = "Y" Then
        isCompleted = True
    Else
        isCompleted = False
    End If
End Property
Public Property Let isCompleted(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isCompleted_) <> Trim(tmpStr) Then
        isCompleted_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get JobDetails() As Collection
Dim tmpS As New BWGSearchManager
    If colJobDet_ Is Nothing Then Set colJobDet_ = tmpS.JobDataDetailSearch("JobID='" & id_ & "'")
    Set tmpS = Nothing
    Set JobDetails = colJobDet_
End Property
Public Property Set JobDetails(ByVal newVal As Collection)
    Set colJobDet_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryJobDetail()
    Set colJobDet_ = Nothing
End Sub

Public Property Get Cars() As Collection
Dim tmpS As New BWGSearchManager
    If colJobCar_ Is Nothing Then Set colJobCar_ = tmpS.JobDataCarSearch("JobID='" & id_ & "'")
    Set tmpS = Nothing
    Set Cars = colJobCar_
End Property
Public Property Set Cars(ByVal newVal As Collection)
    Set colJobCar_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryCar()
    Set colJobCar_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%, isNew As Boolean
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbJobDataHeader Where JobID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    isNew = False
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("JobID") = id_
        isNew = True
    End If
    tmpRec("CustomerID") = Trim(customerID_)
    tmpRec("CustomerSiteID") = Trim(customerSiteID_)
    If (isNew = False And Trim("" & tmpRec("WorkDate")) <> Trim(workDate_)) Or isNew = True Then
        tmpRec("WorkDate") = Trim(workDate_)
        'DBManager.GetDBConnection.Execute "Update tbJobDataTimeTable set WorkDate='" & workDate_ & "' Where JobDataCarID in (Select JobDataCarID From tbJobDataCarUsed Where JobID='" & id_ & "')"
    End If
    tmpRec("WorkTime") = Trim(workTime_)
    tmpRec("IssuedStaffID") = Trim(issuedStaffID_)
    tmpRec("SaleStaffName") = Trim(saleStaffName_)
    tmpRec("RemDesc") = Trim(remDesc_)
    tmpRec("IssuedDate") = Trim(issuedDate_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("isCancel") = Trim(isCancel_)
    tmpRec("WasteCreateNo") = Trim(wasteCreateNo_)
    tmpRec("WasteType") = Trim(wasteType_)
    tmpRec("UnderSaleStaffID") = Trim(underStaffID_)
    If Trim(isCompleted_) = "" Then isCompleted_ = "N"
    tmpRec("isCompleted") = Trim(isCompleted_)
    tmpRec.Update
    tmpRec.Requery
    jobNo_ = Trim("" & tmpRec("JobNo"))
    tmpRec.Close
    With DBManager.GetDBConnection
        If Not colJobDet_ Is Nothing Then
            .Execute "Update tbJobDataDetail Set MarkForDelete='Y' Where JobID='" & id_ & "'"
            For x = 1 To colJobDet_.Count
                colJobDet_(x).Save
            Next
            .Execute "Delete tbJobDataDetail Where MarkForDelete='Y' And JobID='" & id_ & "'"
        End If
        If Not colJobCar_ Is Nothing Then
            .Execute "Update tbJobDataCarUsed Set MarkForDelete='Y' Where JobID='" & id_ & "'"
            For x = 1 To colJobCar_.Count
                colJobCar_(x).Save Trim(workDate_)
            Next
            .Execute "Delete tbJobDataTimeTable Where JobDataCarID in (Select JobDataCarID From tbJobDataCarUsed Where MarkForDelete='Y' And JobID='" & id_ & "')"
            .Execute "Delete tbJobDataCarUsed Where MarkForDelete='Y' And JobID='" & id_ & "'"
        End If
    End With
    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBManager.GetDBConnection.Execute "Delete tbJobDataHeader Where JobID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("JobID"))
    jobNo_ = Trim("" & selRec("JobNo"))
    customerID_ = Trim("" & selRec("CustomerID"))
    customerName_ = Trim("" & selRec("CompanyName"))
    customerSiteID_ = Trim("" & selRec("CustomerSiteID"))
    customerSiteName_ = Trim("" & selRec("Addr1") & " " & selRec("TumbolName") & " " & selRec("AumphurName") & " " & selRec("ProvinceName"))
    customerSiteTel_ = Trim("" & selRec("TelNo"))
    customerSiteFax_ = Trim("" & selRec("FaxNo"))
    customerSiteStaff_ = Trim("" & selRec("ContactName"))
    customerSiteAumphurID_ = Trim("" & selRec("CustSiteAumphurID"))
    customerSiteProvinceID_ = Trim("" & selRec("CustSiteProvinceID"))
    customerSiteProvinceName_ = Trim("" & selRec("ProvinceName"))
    workDate_ = Trim("" & selRec("WorkDate"))
    workTime_ = Trim("" & selRec("WorkTime"))
    issuedStaffID_ = Trim("" & selRec("IssuedStaffID"))
    issuedStaffName_ = Trim("" & selRec("IssuedStaffName"))
    saleStaffName_ = Trim("" & selRec("SaleStaffName"))
    remDesc_ = Trim("" & selRec("RemDesc"))
    issuedDate_ = Trim("" & selRec("IssuedDate"))
    companyID_ = Trim("" & selRec("CompanyID"))
    isCancel_ = Trim("" & selRec("isCancel"))
    underStaffID_ = Trim("" & selRec("UnderSaleStaffID"))
    underStaffName_ = Trim("" & selRec("StaffFName")) & " " & Trim("" & selRec("StaffLName"))
    If IsNull(selRec("MenifestCount")) = False Then menifestCountIssued_ = Trim("" & selRec("MenifestCount"))
    If IsNull(selRec("MenifestCountAll")) = False Then menifestCountAll_ = Trim("" & selRec("MenifestCountAll"))
    wasteCreateNo_ = Trim("" & selRec("WasteCreateNo"))
    wasteType_ = Trim("" & selRec("WasteType"))
    isCompleted_ = Trim("" & selRec("isCompleted"))
End Sub

Public Sub CancelJob()
    With DBManager.GetDBConnection
        .Execute "Update tbJobDataHeader Set isCancel='Y' Where JobID='" & id_ & "'"
        '.Execute "Update tbMenifestHeader Set isCancel='Y' Where CreatedFromJobID='" & id_ & "'"
        isCancel_ = "Y"
    End With
End Sub

Public Sub CloseJob()
    With DBManager.GetDBConnection
        .Execute "Update tbJobDataHeader Set isCompleted='Y' Where JobID='" & id_ & "'"
        isCompleted_ = "Y"
    End With
End Sub

Public Function DuplicateJob(newWorkDate As String) As BWGJobData
Dim newJob As New BWGJobData
Dim x%, newJobDet As BWGJobDataDetail, y%
Dim newJobCar As BWGJobDataCarUsed
Dim newTimeTable As BWGJobDataTimeTable
Dim tmpS As New BWGSearchManager
Dim selDate As String
    If Trim(newWorkDate) = "" Then
        selDate = workDate_
    Else
        selDate = newWorkDate
    End If
    With newJob
        .ID = GetGUID
        .JobNo = jobNo_
        .CustomerID = customerID_
        .CustomerName = customerName_
        .CustomerSiteID = customerSiteID_
        .CustomerSiteName = customerSiteName_
        .CustomerSiteTel = customerSiteTel_
        .CustomerSiteFax = customerSiteFax_
        .CustomerSiteStaff = customerSiteStaff_
        .CustomerSiteAumphurID = customerSiteAumphurID_
        .UnderStaffID = underStaffID_
        .UnderStaffName = underStaffName_
        .WorkDate = selDate
        .WorkTime = workTime_
        .WasteCreateNo = wasteCreateNo_
        .IssuedStaffID = issuedStaffID_
        .IssuedStaffName = issuedStaffName_
        .SaleStaffName = saleStaffName_
        .RemDesc = remDesc_
        .IssuedDate = issuedDate_
        .companyID = companyID_
        If isCancel_ = "Y" Then
            .isCancel = True
        Else
            .isCancel = False
        End If
        If colJobDet_ Is Nothing Then Set colJobDet_ = tmpS.JobDataDetailSearch("JobID='" & id_ & "'")
        If Not colJobDet_ Is Nothing Then
            For x = 1 To colJobDet_.Count
                Set newJobDet = New BWGJobDataDetail
                With colJobDet_(x)
                    newJobDet.ContainerCount = .ContainerCount
                    newJobDet.ContainerTypeDesc = .ContainerTypeDesc
                    newJobDet.DestinationDesc = .DestinationDesc
                    newJobDet.EstWasteQty = 0 '.EstWasteQty
                    newJobDet.ID = GetGUID
                    newJobDet.JobID = newJob.ID
                    newJobDet.isPriceIncTrans = .isPriceIncTrans
                    newJobDet.JobNo = ""
                    newJobDet.RemDesc = .RemDesc
                    newJobDet.rowID = .rowID
                    newJobDet.WasteDataID = .WasteDataID
                    newJobDet.WasteEUCode = .WasteEUCode
                    newJobDet.WasteName = .WasteName
                    newJobDet.WasteNo = .WasteNo
                    newJobDet.WasteQtyUnitID = .WasteQtyUnitID
                    newJobDet.WasteQtyUnitName = .WasteQtyUnitName
                    newJobDet.WasteQuotationNo = .WasteQuotationNo
                    newJobDet.TSDFID = .TSDFID
                    newJobDet.TSDFName = .TSDFName
                    newJobDet.TSDFRegisNo = .TSDFRegisNo
                    newJobDet.TSDFTumbolID = .TSDFTumbolID
                    newJobDet.carNo = .carNo
                    newJobDet.QuoLaborCharge = .QuoLaborCharge
                    newJobDet.QuoLaborCount = .QuoLaborCount
                    newJobDet.QuotationID = .QuotationID
                    newJobDet.QuoTransportFee = .QuoTransportFee
                    newJobDet.QuoTransUnitID = .QuoTransUnitID
                    newJobDet.QuoTransUnitName = .QuoTransUnitName
                    newJobDet.QuoTreatmentRate = .QuoTreatmentRate
                    newJobDet.QuoTreatmentUnitID = .QuoTreatmentUnitID
                    newJobDet.QuoTreatmentUnitName = .QuoTreatmentUnitName
                    newJobDet.QuoTruckTypeID = .QuoTruckTypeID
                    newJobDet.QuoTruckTypeName = .QuoTruckTypeName
                    newJobDet.QuoMinWeightPerCar = .QuoMinWeightPerCar
                    newJobDet.NonMinRate = .NonMinRate
                    newJobDet.HazMinRate = .HazMinRate
                    newJobDet.NBLMinRate = .NBLMinRate
                    newJobDet.HazMinRate = .HazMinRate
                End With
                .JobDetails.Add newJobDet, "'" & newJobDet.ID & "'"
                Set newJobDet = Nothing
            Next
        End If
        If colJobCar_ Is Nothing Then Set colJobCar_ = tmpS.JobDataCarSearch("JobID='" & id_ & "'")
        If Not colJobCar_ Is Nothing Then
            For x = 1 To colJobCar_.Count
                Set newJobCar = New BWGJobDataCarUsed
                With colJobCar_(x)
                    newJobCar.DestinationDesc = .DestinationDesc
                    newJobCar.ID = GetGUID
                    newJobCar.JobID = newJob.ID
                    newJobCar.Remark = .Remark
                    newJobCar.rowID = .rowID
                    newJobCar.TruckCount = .TruckCount
                    newJobCar.isTrailer = .isTrailer
                    newJobCar.TruckSubTypeID = .TruckSubTypeID
                    newJobCar.TruckSubTypeName = .TruckSubTypeName
                    newJobCar.TruckTypeName = .TruckTypeName
                    newJobCar.WorkerPerCar = .WorkerPerCar
                    newJobCar.GoSameTime = .GoSameTime
                    newJobCar.TakePhoto = .TakePhoto
                    newJobCar.carNo = .carNo
                    
                    For y = 1 To .TimeTables.Count
                        Set newTimeTable = New BWGJobDataTimeTable
                        newTimeTable.JobDataCarID = newJobCar.ID
                        newTimeTable.StartTime = .TimeTables(y).StartTime
                        newTimeTable.tripNo = .TimeTables(y).tripNo
                        newTimeTable.WorkDate = selDate
                        newTimeTable.WorkerCount = .WorkerPerCar
                        newJobCar.TimeTables.Add newTimeTable, "'" & newTimeTable.JobDataCarID & "_" & newTimeTable.tripNo & "'"
                        Set newTimeTable = Nothing
                    Next
                End With
                .Cars.Add newJobCar, "'" & newJobCar.ID & "'"
               
                Set newJobCar = Nothing
            Next
        End If
    End With
    Set DuplicateJob = newJob
    Set newJob = Nothing
    Set tmpS = Nothing
End Function

Public Function GenerateMenifest() As Integer
Dim tmpcar As BWGJobDataCarUsed, tmpTimeTable As BWGJobDataTimeTable
Dim tmpMenifest As BWGMenifest, tmpCon As ADODB.Connection
Dim x%, y%, iCount%, z%, colWaste As Collection, aa%
Dim tmpCol As Collection, tmpS As New BWGSearchManager
Dim enumWType(1 To 4) As String, colGenNo As Collection, bb%
    Set colGenNo = tmpS.WasteGenNoSearch("JobID='" & id_ & "'")
    enumWType(1) = "Non"
    enumWType(2) = "Haz"
    enumWType(3) = "NBL"
    enumWType(4) = "HBL"
    Set tmpCon = DBManager.GetDBConnection
            
    For x = 1 To colJobCar_.Count
        Set tmpcar = colJobCar_(x)
        For aa = 1 To 4
            For y = 1 To tmpcar.TimeTables.Count
                Set tmpTimeTable = tmpcar.TimeTables(y)
                For bb = 1 To colGenNo.Count
                    Set colWaste = GetJobDetByWasteType(aa, tmpcar.carNo, tmpTimeTable.JobDetails, colGenNo(bb).WasteCreateNo)
                    If colWaste.Count > 0 Then
                        If Trim(tmpTimeTable.CarRegisNo) <> "" Then
                            'If Trim(tmpTimeTable.MenifestID) = "" Then
                            '    tmpTimeTable.MenifestID = GetGUID
                            '    tmpCon.Execute "Update tbJobDataTimeTable Set MenifestID='" & tmpTimeTable.MenifestID & "' Where JobDataCarID='" & tmpTimeTable.JobDataCarID & "' And TripNo=" & tmpTimeTable.TripNo
                            'End If
                            Set tmpMenifest = New BWGMenifest
                            With tmpMenifest
                                .ID = GetGUID
                                .CreatedFromJobID = id_
                                .companyID = companyID_
                                .CustomerID = customerID_
                                .CustSiteID = customerSiteID_
                                .WasteGenNo = colGenNo(bb).WasteCreateNo
                                .NonOrHz = enumWType(aa)
                                .TelNo = customerSiteTel_
                                .FaxNo = customerSiteFax_
                                .ChargeWType = "A"
                                .ErContact = customerSiteStaff_
                                If Trim(tmpcar.TruckTypeName) <> "รถโรงงาน" Then
                                    .TransportName1 = "เบตเตอร์ เวิลด์ ทรานสปอร์ตจำกัด"
                                    .TransRegisNo1 = "DIW-T-050200740"
                                Else
                                    .TransportName1 = ""
                                    .TransRegisNo1 = ""
                                End If
                                .Transport2ID = tmpTimeTable.VendorID
                                .TransportName2 = tmpTimeTable.VendorName
                                .TransRegisNo2 = tmpTimeTable.VendorRegisNo
                                .OutWorkerFee = tmpTimeTable.VendorWorkerCharge
                                .TransTelNo1 = ""
                                .TransFaxNo1 = ""
                                .TransErContact1 = ""
                                .TruckSubTypeID = tmpcar.TruckSubTypeID
                                .CarRegisID = tmpTimeTable.CarRegisNo
                                .OutWorkerCount = tmpcar.WorkerPerCar
                                .FromProvinceID = customerSiteProvinceID_
                                .FromProvinceName = customerSiteProvinceName_
                                .OpenStaffID = CurrentUser.ID
                                .OpenDate = TodayDate
                                .WorkDate = workDate_
                                .isClosed = False
                                .isCanceled = False
                                .JobDataCarID = tmpTimeTable.JobDataCarID
                                .tripNo = tmpTimeTable.tripNo
                                .DriverName = tmpTimeTable.DriverName
                                .UnderStaffID = underStaffID_
                                .UnderStaffName = underStaffName_
                                For z = 1 To colWaste.Count
                                    If colWaste(z).carNo = tmpcar.carNo And colJobDet_(z).MenifestID = "" Then
                                        tmpTimeTable.JobDetails.Add colWaste(z), "'" & colWaste(z).ID & "'"
                                        .JobDetails.Add colWaste(z), "'" & colWaste(z).ID & "'"
                                        If Trim(.TSDF1ID) = "" Then
                                            Set tmpCol = tmpS.TSDFSearch("TSDFID='" & colWaste(z).TSDFID & "'")
                                            If tmpCol.Count > 0 Then
                                                .TSDFName1 = tmpCol(1).CompanyName
                                                .TSDF1ID = tmpCol(1).ID
                                                .TSDFAddr = tmpCol(1).Addr1
                                                .TSDFAumphurID = tmpCol(1).AumphurID
                                                .TSDFAumphurName = tmpCol(1).AumphurName
                                                .TSDFErContact = tmpCol(1).ContactName
                                                .TSDFFaxNo = tmpCol(1).FaxNo
                                                .TSDFProvinceName = tmpCol(1).ProvinceName
                                                .TSDFRegisNo1 = tmpCol(1).TSDFRegisNo
                                                .TSDFTelNo = tmpCol(1).TelNo
                                                .TSDFTumbolID = tmpCol(1).TumbolID
                                                .TSDFTumbolName = tmpCol(1).TumbolName
                                                .ToProvinceID = tmpCol(1).ProvinceID
                                                .ToProvinceName = tmpCol(1).ProvinceName
                                            End If
                                            Set tmpCol = Nothing
                                            Set tmpS = Nothing
                                        End If
                                    End If
                                Next
                                .Save
                                iCount = iCount + 1
                            End With
                        End If
                    End If
                Next bb
            Next
        Next
    Next
    Set tmpCon = Nothing
    GenerateMenifest = iCount
    RequeryMenifestCountIssued
End Function

Public Sub RequeryMenifestCountIssued()
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
    tmpRec.Open "Select * from vw_MenifestCountByJobHeader Where JobID='" & id_ & "'", DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    If tmpRec.EOF = False Then
        menifestCountIssued_ = tmpRec("MenifestCount")
    End If
    tmpRec.Close
    Set tmpRec = Nothing
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub

Private Function GetJobDetByWasteType(selType As Integer, carNo As Integer, selCol As Collection, genNoID As String) As Collection
Dim wType As String
Dim x%, tmpCol As Collection
Dim aa As BWGJobDataDetail
    
    Select Case selType
        Case 1
            wType = "Non"
        Case 2
            wType = "Haz"
        Case 3
            wType = "NBL"
        Case 4
            wType = "HBL"
    End Select
    Set tmpCol = New Collection
    For x = 1 To colJobDet_.Count
        If colJobDet_(x).WasteType = wType And colJobDet_(x).WasteGenNo = genNoID And colJobDet_(x).carNo = carNo And isExist(selCol, "'" & colJobDet_(x).ID & "'") = False Then
            tmpCol.Add colJobDet_(x), "'" & colJobDet_(x).ID & "'"
        End If
    Next
    Set GetJobDetByWasteType = tmpCol
    Set tmpCol = Nothing
End Function
