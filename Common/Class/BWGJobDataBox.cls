VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGJobDataBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim jobNo_ As String
Dim CustomerID_ As String
Dim CustomerName_ As String
Dim CustomerSiteID_ As String
Dim CustomerSiteName_ As String
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
Dim SaleStaffName_ As String
Dim remDesc_ As String
Dim issuedDate_ As String
Dim IssuedTime_ As String
Dim companyID_ As String
Dim isCancel_ As String
Dim wasteCreateNo_ As String
Dim wasteType_ As String
Dim colJobDet_ As Collection
Dim colJobCar_ As Collection
Dim menifestCountAll_ As Integer
Dim menifestCountIssued_ As Integer
Dim saleDistrictID_ As String
Dim underStaffID_ As String
Dim underStaffName_ As String
Dim isCompleted_ As String
Dim isChanged_ As Boolean
Dim isJob_ As String
Dim isCopy_ As String
Dim isNotActive_ As String
Dim JobType_ As String

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
    CustomerID = CustomerID_
End Property
Public Property Let CustomerID(ByVal newVal As String)
    If Trim(CustomerID_) <> Trim(newVal) Then
        CustomerID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerName() As String
    CustomerName = CustomerName_
End Property
Public Property Let CustomerName(ByVal newVal As String)
    If Trim(CustomerName_) <> Trim(newVal) Then
        CustomerName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteID() As String
    CustomerSiteID = CustomerSiteID_
End Property
Public Property Let CustomerSiteID(ByVal newVal As String)
    If Trim(CustomerSiteID_) <> Trim(newVal) Then
        CustomerSiteID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteName() As String
    CustomerSiteName = CustomerSiteName_
End Property
Public Property Let CustomerSiteName(ByVal newVal As String)
    If Trim(CustomerSiteName_) <> Trim(newVal) Then
        CustomerSiteName_ = Trim(newVal)
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
    SaleStaffName = SaleStaffName_
End Property
Public Property Let SaleStaffName(ByVal newVal As String)
    If Trim(SaleStaffName_) <> Trim(newVal) Then
        SaleStaffName_ = Trim(newVal)
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

Public Property Get IssuedTime() As String
    IssuedTime = IssuedTime_
End Property
Public Property Let IssuedTime(ByVal newVal As String)
    If Trim(IssuedTime_) <> Trim(newVal) Then
        IssuedTime_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyID() As String
    CompanyID = companyID_
End Property
Public Property Let CompanyID(ByVal newVal As String)
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

'Public Property Get MenifestCountAll() As Integer
'    MenifestCountAll = menifestCountAll_
'End Property
'Public Property Let MenifestCountAll(ByVal newVal As Integer)
'    If menifestCountAll_ <> newVal Then
'        menifestCountAll_ = newVal
'        isChanged_ = True
'    End If
'End Property
'
'Public Property Get MenifestCountIssued() As Integer
'    MenifestCountIssued = menifestCountIssued_
'End Property
'Public Property Let MenifestCountIssued(ByVal newVal As Integer)
'    If menifestCountIssued_ <> newVal Then
'        menifestCountIssued_ = newVal
'        isChanged_ = True
'    End If
'End Property

Public Property Get SaleDistrictID() As String
    SaleDistrictID = saleDistrictID_
End Property
Public Property Let SaleDistrictID(ByVal newVal As String)
    If Trim(saleDistrictID_) <> newVal Then
        saleDistrictID_ = Trim(newVal)
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

'Public Property Get JobDetails() As Collection
'Dim tmpS As New BWGSearchManager
'    If colJobDet_ Is Nothing Then Set colJobDet_ = tmpS.JobDataDetailSearch("JobID='" & id_ & "'")
'    Set tmpS = Nothing
'    Set JobDetails = colJobDet_
'End Property
'Public Property Set JobDetails(ByVal newVal As Collection)
'    Set colJobDet_ = newVal
'    isChanged_ = True
'End Property
'
'Public Sub RequeryJobDetail()
'    Set colJobDet_ = Nothing
'End Sub

Public Property Get Cars() As Collection
Dim tmpS As New BWGSearchManager
    If colJobCar_ Is Nothing Then Set colJobCar_ = tmpS.JobDataCarBoxSearch("JobID='" & id_ & "'")
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

Public Property Get isJob() As String
    isJob = isJob_
End Property
Public Property Let isJob(ByVal newVal As String)
    If Trim(isJob_) <> newVal Then
        isJob_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isCopy() As String
    isCopy = isCopy_
End Property
Public Property Let isCopy(ByVal newVal As String)
    If Trim(isCopy_) <> newVal Then
        isCopy_ = Trim(newVal)
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Property Get isNotActive() As String
    isNotActive = isNotActive_
End Property
Public Property Let isNotActive(ByVal newVal As String)
     If Trim(isNotActive_) <> newVal Then
        isNotActive_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get JobType() As String
    JobType = JobType_
End Property
Public Property Let JobType(ByVal newVal As String)
    If Trim(JobType_) <> newVal Then
        JobType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property


Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset, RS As ADODB.Recordset
Dim X%, isNew As Boolean
Dim sqlStr As String
Dim LastUpdate As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbJobDataHeaderBox WHERE JobID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    isNew = False
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("JobID") = id_
        isNew = True
    End If
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("CustomerSiteID") = Trim(CustomerSiteID_)
'    If (isNew = False And Trim("" & tmpRec("WorkDate")) <> Trim(workDate_)) Or isNew = True Then
        tmpRec("WorkDate") = Trim(workDate_)
        'DBConnExc "Update tbJobDataTimeTableBox set WorkDate='" & workDate_ & "' WHERE JobDataCarID in (Select JobDataCarID From tbJobDataCarUsedBox WHERE JobID='" & id_ & "')"
'    End If
    tmpRec("WorkTime") = Trim(workTime_)
    tmpRec("IssuedStaffID") = Trim(issuedStaffID_)
    tmpRec("UnderSaleStaffID") = Trim(underStaffID_)
    tmpRec("SaleStaffName") = Trim(SaleStaffName_)
    tmpRec("RemDesc") = Trim(remDesc_)
    tmpRec("IssuedDate") = Trim(issuedDate_)
    tmpRec("IssuedTime") = Trim(IssuedTime_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("isCancel") = IIf(Trim(isCancel_) = "", "N", Trim(isCancel_))
    tmpRec("WasteCreateNo") = Trim(wasteCreateNo_)
    tmpRec("WasteType") = Trim(wasteType_)
    If Trim(isCompleted_) = "" Then isCompleted_ = "N"
    tmpRec("isCompleted") = Trim(isCompleted_)
    tmpRec!isJobs = IIf(isJob_ = "", "N", isJob_)
    tmpRec!isNotActive = IIf(isNotActive_ = "", "N", isNotActive_)
    tmpRec("JobType") = Trim(JobType_)
'    tmpRec!LastUserUpdate = Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName)
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    
    tmpRec.Update
    tmpRec.Requery
    jobNo_ = Trim("" & tmpRec("JobNo"))
    
'    DBConn.Execute "UPDATE tbJobDataHeaderBox SET LastUpdate = CONVERT(DATETIME, GETDATE(), 103) WHERE JobID='" & id_ & "'"
    
'    tmpRec.Close

'        If Not colJobDet_ Is Nothing Then
'            DBConnExc "Update tbJobDataDetail Set MarkForDelete='Y' WHERE JobID='" & id_ & "'"
'            For x = 1 To colJobDet_.Count
'                colJobDet_(x).RowID = x
'                colJobDet_(x).Save
'            Next
'            DBConnExc "Delete tbJobDataDetail WHERE MarkForDelete='Y' And JobID='" & id_ & "'"
'        End If
        If Not colJobCar_ Is Nothing Then
            DBConnExc "Update tbJobDataCarUsedBox Set MarkForDelete='Y' WHERE JobID='" & id_ & "'"
            For X = 1 To colJobCar_.Count
                colJobCar_(X).Save Trim(workDate_), isCopy_
            Next
            DBConnExc "Delete tbJobDataTimeTableBox WHERE JobDataCarID in (Select JobDataCarID From tbJobDataCarUsedBox WHERE MarkForDelete='Y' And JobID='" & id_ & "')"
            DBConnExc "Delete tbJobDataCarUsedBox WHERE MarkForDelete='Y' And JobID='" & id_ & "'"
        End If
        
    Set tmpRec = Nothing
    isChanged_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbJobDataHeaderBox WHERE JobID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("JobID"))
    jobNo_ = Trim("" & selRec("JobNo"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    CustomerName_ = Trim("" & selRec("CompanyName"))
    CustomerSiteID_ = Trim("" & selRec("CustomerSiteID"))
    CustomerSiteName_ = Trim("" & selRec("Addr1") & " " & selRec("TumbolName") & " " & selRec("AumphurName") & " " & selRec("ProvinceName"))
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
    SaleStaffName_ = Trim("" & selRec("SaleStaffName"))
    remDesc_ = Trim("" & selRec("RemDesc"))
    issuedDate_ = Trim("" & selRec("IssuedDate"))
    IssuedTime_ = Trim("" & selRec("IssuedTime"))
    companyID_ = Trim("" & selRec("CompanyID"))
    isCancel_ = Trim("" & selRec("isCancel"))
    If isExist(selRec, "SaleDistrictID") Then saleDistrictID_ = Trim("" & selRec("SaleDistrictID"))
    underStaffID_ = Trim("" & selRec("UnderSaleStaffID"))
    underStaffName_ = Trim("" & selRec("StaffFName")) & " " & Trim("" & selRec("StaffLName"))
'    If IsNull(selRec("MenifestCount")) = False Then menifestCountIssued_ = Trim("" & selRec("MenifestCount"))
'    If IsNull(selRec("MenifestCountAll")) = False Then menifestCountAll_ = Trim("" & selRec("MenifestCountAll"))
    wasteCreateNo_ = Trim("" & selRec("WasteCreateNo"))
    wasteType_ = Trim("" & selRec("WasteType"))
    isCompleted_ = Trim("" & selRec("isCompleted"))
    isJob_ = selRec!isJobs
    isNotActive_ = selRec!isNotActive
    JobType_ = selRec!JobType
End Sub

Public Sub CancelJob()
    Dim tSql As String
    DBConnExc "Update tbJobDataHeaderBox Set isCancel='Y' WHERE JobID='" & id_ & "'"
    isCancel_ = "Y"
    
'    tSql = "IF  EXISTS (SELECT LogJobNo FROM Box_LogMessage WHERE LogJobNo = '" & jobNo_ & "' AND LogState = 1)" & vbCrLf
'    tSql = tSql & "DELETE Box_LogMessage WHERE LogJobNo = '" & jobNo_ & "' AND LogState = 1"
'    GetDBConnection DBConStock
'    DBConn.Execute tSql
    GetDBConnection
End Sub

Public Sub CloseJob()
    DBConnExc "Update tbJobDataHeaderBox Set isCompleted='Y' WHERE JobID='" & id_ & "'"
    isCompleted_ = "Y"
End Sub

'Public Function DuplicateJob(newWorkDate As String) As BWGJobDataBox
'Dim newJob As New BWGJobDataBox
'Dim x%, newJobDet As BWGJobDataDetail, y%
'Dim newJobCar As BWGJobDataCarUsedBox
'Dim newTimeTable As BWGJobDataTimeTable
'Dim tmpS As New BWGSearchManager
'Dim colCustomer As Collection, JobCustomer As New BWGCustomer
'Dim selDate As String
'    If Trim(newWorkDate) = "" Then
'        selDate = workDate_
'    Else
'        selDate = newWorkDate
'    End If
'    With newJob
'        .ID = GetGUID
'        .JobNo = jobNo_
'        .CustomerID = CustomerID_
'        .CustomerName = customerName_
'        .CustomerSiteID = CustomerSiteID_
'        .CustomerSiteName = CustomerSiteName_
'        .CustomerSiteTel = customerSiteTel_
'        .CustomerSiteFax = customerSiteFax_
'        .CustomerSiteStaff = customerSiteStaff_
'        .CustomerSiteAumphurID = customerSiteAumphurID_
'        Set colCustomer = tmpS.CustomerSearch("CustomerID = '" & CustomerID_ & "'")
'        Set JobCustomer = colCustomer(1)
'        .UnderStaffID = JobCustomer.UnderSaleStaffID
'        .UnderStaffName = JobCustomer.UnderSaleName
'        Set JobCustomer = Nothing
'        Set colCustomer = Nothing
'        .WorkDate = selDate
'        .WorkTime = workTime_
'        .WasteCreateNo = wasteCreateNo_
'        .IssuedStaffID = issuedStaffID_
'        .IssuedStaffName = issuedStaffName_
'        .SaleStaffName = saleStaffName_
'        .RemDesc = remDesc_
'        .IssuedDate = issuedDate_
'        .IssuedTime = IssuedTime_
'        .CompanyID = companyID_
'        If isCancel_ = "Y" Then
'            .isCancel = True
'        Else
'            .isCancel = False
'        End If
'
'        If colJobDet_ Is Nothing Then Set colJobDet_ = tmpS.JobDataDetailSearch("JobID='" & id_ & "'")
'        If Not colJobDet_ Is Nothing Then
'            For x = 1 To colJobDet_.Count
'                Set newJobDet = New BWGJobDataBoxDetail
'                With colJobDet_(x)
'                    newJobDet.ContainerCount = .ContainerCount
'                    newJobDet.ContainerTypeDesc = .ContainerTypeDesc
'                    newJobDet.DestinationDesc = .DestinationDesc
'                    newJobDet.EstWasteQty = 0 '.EstWasteQty
'                    newJobDet.ID = GetGUID
'                    newJobDet.JobID = newJob.ID
'                    newJobDet.isPriceIncTrans = .isPriceIncTrans
'                    newJobDet.JobNo = ""
'                    newJobDet.RemDesc = .RemDesc
'                    newJobDet.RowID = .RowID
'                    newJobDet.wastedataID = .wastedataID
'                    newJobDet.WasteEUCode = .WasteEUCode
'                    newJobDet.WasteName = .WasteName
'                    newJobDet.WasteNo = .WasteNo
'                    newJobDet.WasteQtyUnitID = .WasteQtyUnitID
'                    newJobDet.WasteQtyUnitName = .WasteQtyUnitName
'                    newJobDet.WasteQuotationNo = .WasteQuotationNo
'                    newJobDet.TSDFID = .TSDFID
'                    newJobDet.TSDFName = .TSDFName
'                    newJobDet.TSDFRegisNo = .TSDFRegisNo
'                    newJobDet.TSDFTumbolID = .TSDFTumbolID
'                    newJobDet.carNo = .carNo
'                    newJobDet.QuoLaborCharge = .QuoLaborCharge
'                    newJobDet.QuoLaborCount = .QuoLaborCount
'                    newJobDet.QuotationID = .QuotationID
'                    newJobDet.QuoTransportFee = .QuoTransportFee
'                    newJobDet.QuoTransUnitID = .QuoTransUnitID
'                    newJobDet.QuoTransUnitName = .QuoTransUnitName
'                    newJobDet.QuoTreatmentRate = .QuoTreatmentRate
'                    newJobDet.QuoTreatmentUnitID = .QuoTreatmentUnitID
'                    newJobDet.QuoTreatmentUnitName = .QuoTreatmentUnitName
'                    newJobDet.QuoTruckTypeID = .QuoTruckTypeID
'                    newJobDet.QuoTruckTypeName = .QuoTruckTypeName
'                    newJobDet.QuoMinWeightPerCar = .QuoMinWeightPerCar
'                    newJobDet.NonMinRate = .NonMinRate
'                    newJobDet.HazMinRate = .HazMinRate
'                    newJobDet.NBLMinRate = .NBLMinRate
'                    newJobDet.HazMinRate = .HazMinRate
'                    ''******************** แก้ไข ข้อมูลราคาและหน่วยไม่ถูกก็อปปี้ ทำให้ไม่คำนวณข้อมูลในการออก Invoice
'                    newJobDet.JobTreatmentRate = .JobTreatmentRate
'                    newJobDet.JobTreatmentUnitID = .JobTreatmentUnitID
'                    newJobDet.JobTransportFee = .JobTransportFee
'                    newJobDet.JobTransUnitID = .JobTransUnitID
'                    ''**************************************************************
'                End With
'                .JobDetails.Add newJobDet, "'" & newJobDet.ID & "'"
'                Set newJobDet = Nothing
'            Next
'        End If
'        If colJobCar_ Is Nothing Then Set colJobCar_ = tmpS.JobDataCarSearch("JobID='" & id_ & "'")
'        If Not colJobCar_ Is Nothing Then
'            For x = 1 To colJobCar_.Count
'                Set newJobCar = New BWGJobDataBoxCarUsed
'                With colJobCar_(x)
'
'                    Dim TruckSubTypeID2 As String
'                    TruckSubTypeID2 = ""
'                    If .TruckSubTypeID = "33D46FC64F3D450897D1F72986EAE700" Then TruckSubTypeID2 = "F0D9A37CFC034EB89BF0F0DBA15D9523"
'                    If .TruckSubTypeID = "F0D9A37CFC034EB89BF0F0DBA15D9523" Then TruckSubTypeID2 = "33D46FC64F3D450897D1F72986EAE700"
'                    If .TruckSubTypeID = "2EFEDFB372C448E3B93C786DA8613F6A" Then TruckSubTypeID2 = "CD50C04FD207477BBA66216557868208"
'                    If .TruckSubTypeID = "CD50C04FD207477BBA66216557868208" Then TruckSubTypeID2 = "2EFEDFB372C448E3B93C786DA8613F6A"
'
'                    TruckAvailable.TruckAvailable = 0
'                    TruckAvailable.TruckTotal = 0
'                    TruckAvailableCount newJob.WorkDate, .TruckSubTypeID, TruckSubTypeID2
'                    If TruckAvailable.TruckTotal > 0 Then
'                        If TruckAvailable.TruckAvailable < .TruckCount Then
'                            MsgBox "จำนวนรถ """ & newJobCar.TruckSubTypeName & """ ไม่พอ ไม่สามารถลงจำนวนรถได้" & vbCrLf & "กรุณาตรวจสอบจำนวนรถของวันนี้กับฝ่ายขนส่งอีกครั้ง, จำนวนรถคงเหลือ (" & TruckAvailable.TruckAvailable & ")", vbInformation
'                            Exit For
'                        End If
'                    End If
'
'                    newJobCar.DestinationDesc = .DestinationDesc
'                    newJobCar.ID = GetGUID
'                    newJobCar.JobID = newJob.ID
'                    newJobCar.Remark = .Remark
'                    newJobCar.RowID = .RowID
'                    newJobCar.TruckCount = .TruckCount
'                    newJobCar.isTrailer = .isTrailer
'                    newJobCar.TruckSubTypeID = .TruckSubTypeID
'                    newJobCar.TruckSubTypeName = .TruckSubTypeName
'                    newJobCar.TruckTypeName = .TruckTypeName
'                    newJobCar.WorkerPerCar = .WorkerPerCar
'                    newJobCar.GoSameTime = .GoSameTime
'                    newJobCar.TakePhoto = .TakePhoto
'                    newJobCar.carNo = .carNo
'
'                    For y = 1 To .TimeTables.Count
'                        Set newTimeTable = New BWGJobDataBoxTimeTable
'                        newTimeTable.JobDataCarID = newJobCar.ID
'                        newTimeTable.StartTime = .TimeTables(y).StartTime
'                        newTimeTable.TripNo = .TimeTables(y).TripNo
'                        newTimeTable.WorkDate = selDate
'                        newTimeTable.WorkerCount = .WorkerPerCar
'                        newJobCar.TimeTables.Add newTimeTable, "'" & newTimeTable.JobDataCarID & "_" & newTimeTable.TripNo & "'"
'                        Set newTimeTable = Nothing
'                    Next
'                End With
'                .Cars.Add newJobCar, "'" & newJobCar.ID & "'"
'
'                Set newJobCar = Nothing
'            Next
'        End If
'    End With
'    Set DuplicateJob = newJob
'    Set newJob = Nothing
'    Set tmpS = Nothing
'End Function

'Public Function GenerateMenifest() As Integer
'Dim tmpCar As BWGJobDataCarUsedBox, tmpTimeTable As BWGJobDataTimeTableBox
'Dim tmpMenifest As BWGMenifest, tmpCon As ADODB.Connection
'Dim x%, y%, iCount%, z%, colWaste As Collection, aa%
'Dim tmpCol As Collection, tmpS As New BWGSearchManager
'Dim enumWType() As String, colGenNo As Collection, bb%
'    Set colGenNo = tmpS.WasteGenNoSearch("JobID='" & id_ & "'")
'    ReDim enumWType(WasteTypes.Count)
'    For x = 1 To WasteTypes.Count
'        enumWType(x) = WasteTypes(x).WasteType
'    Next
'
'    For x = 1 To colJobCar_.Count
'        Set tmpCar = colJobCar_(x)
'        For aa = 1 To WasteTypes.Count
'            For y = 1 To tmpCar.TimeTables.Count
'                Set tmpTimeTable = tmpCar.TimeTables(y)
'                For bb = 1 To colGenNo.Count
'                    Set colWaste = GetJobDetByWasteType(aa, tmpCar.carNo, tmpTimeTable.JobDetails, colGenNo(bb).WasteCreateNo)
'                    If colWaste.Count > 0 Then
'                        If Trim(tmpTimeTable.CarRegisNo) <> "" Then
'                            'If Trim(tmpTimeTable.MenifestID) = "" Then
'                            '    tmpTimeTable.MenifestID = GetGUID
'                            '    DBConnExc "Update tbJobDataTimeTableBox Set MenifestID='" & tmpTimeTable.MenifestID & "' WHERE JobDataCarID='" & tmpTimeTable.JobDataCarID & "' And TripNo=" & tmpTimeTable.TripNo
'                            'End If
'                            Set tmpMenifest = New BWGMenifest
'                            With tmpMenifest
'                                .ID = GetGUID
'                                .CreatedFromJobID = id_
'                                .CompanyID = companyID_
'                                .CustomerID = CustomerID_
'                                .CustSiteID = CustomerSiteID_
'                                .WasteGenNo = colGenNo(bb).WasteCreateNo
'                                .NonOrHz = enumWType(aa)
'                                .TelNo = customerSiteTel_
'                                .FaxNo = customerSiteFax_
'
'                                '*** ตรวจสอบการคิดน้ำหนักจากผู้รับกำจัด หรือผู้ก่อนกำเนิด
'                                Dim tSql As String
'                                Dim RS As New ADODB.Recordset
'                                tSql = "SELECT * FROM tbCustomer_TR_ChargeWType WHERE CustomerID = '" & CustomerID_ & "'"
'                                Set RS = New ADODB.Recordset
'                                Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
'                                If RS.RecordCount > 0 Then
'                                    Select Case RS!ChargeWType
'                                        Case "A"
'                                            .ChargeWType = "A"
'                                        Case "B"
'                                            .ChargeWType = "B"
'                                        Case Else
'                                            .ChargeWType = "A"
'                                    End Select
'                                Else
'                                    .ChargeWType = "A"
'                                End If
'                                '*************************
'
'                                .ErContact = customerSiteStaff_
'                                If Trim(tmpCar.TruckTypeName) <> "รถโรงงานลูกค้า" Then
'                                    .TransportName1 = "เบตเตอร์ เวิลด์ ทรานสปอร์ตจำกัด"
'                                    .TransRegisNo1 = "DIW-T-050200740"
'                                Else
'                                    .TransportName1 = ""
'                                    .TransRegisNo1 = ""
'                                End If
'                                .Transport2ID = tmpTimeTable.VendorID
'                                .TransportName2 = tmpTimeTable.VendorName
'                                .TransRegisNo2 = tmpTimeTable.VendorRegisNo
'                                .OutWorkerFee = tmpTimeTable.VendorWorkerCharge
'                                .TransTelNo1 = ""
'                                .TransFaxNo1 = ""
'                                .TransErContact1 = ""
'                                .TruckSubTypeID = tmpCar.TruckSubTypeID
'                                .CarRegisID = tmpTimeTable.CarRegisNo
'                                .OutWorkerCount = tmpCar.WorkerPerCar
'                                .FromProvinceID = customerSiteProvinceID_
'                                .FromProvinceName = customerSiteProvinceName_
'                                .OpenStaffID = CurrentUser.ID
'                                .OpenDate = TodayDate
'                                .WorkDate = workDate_
'                                .isClosed = False
'                                .IsCanceled = False
'                                .JobDataCarID = tmpTimeTable.JobDataCarID
'                                .TripNo = tmpTimeTable.TripNo
'                                .DriverName = tmpTimeTable.DriverName
'                                .SaleDistrictID = saleDistrictID_
'                                .UnderStaffID = underStaffID_
'                                .UnderStaffName = underStaffName_
'                                For z = 1 To colWaste.Count
'                                    If colWaste(z).carNo = tmpCar.carNo And colJobDet_(z).menifestID = "" Then
'                                        tmpTimeTable.JobDetails.Add colWaste(z), "'" & colWaste(z).ID & "'"
'                                        .JobDetails.Add colWaste(z), "'" & colWaste(z).ID & "'"
'                                        If Trim(.TSDF1ID) = "" Then
'                                            Set tmpCol = tmpS.TSDFSearch("TSDFID='" & colWaste(z).TSDFID & "'")
'                                            If tmpCol.Count > 0 Then
'                                                .TSDFName1 = tmpCol(1).CompanyName
'                                                .TSDF1ID = tmpCol(1).ID
'                                                .TSDFAddr = tmpCol(1).Addr1
'                                                .TSDFAumphurID = tmpCol(1).AumphurID
'                                                .TSDFAumphurName = tmpCol(1).AumphurName
'                                                .TSDFErContact = tmpCol(1).ContactName
'                                                .TSDFFaxNo = tmpCol(1).FaxNo
'                                                .TSDFProvinceName = tmpCol(1).ProvinceName
'                                                .TSDFRegisNo1 = tmpCol(1).TSDFRegisNo
'                                                .TSDFTelNo = tmpCol(1).TelNo
'                                                .TSDFTumbolID = tmpCol(1).TumbolID
'                                                .TSDFTumbolName = tmpCol(1).TumbolName
'                                                .ToProvinceID = tmpCol(1).ProvinceID
'                                                .ToProvinceName = tmpCol(1).ProvinceName
'                                            End If
'                                            Set tmpCol = Nothing
'                                            Set tmpS = Nothing
'                                        End If
'                                    End If
'                                Next
'                                .isJob = isJob_
'                                .Save
'                                iCount = iCount + 1
'                            End With
'                        End If
'                    End If
'                Next bb
'            Next
'        Next
'    Next
'    Set tmpCon = Nothing
'    GenerateMenifest = iCount
'    RequeryMenifestCountIssued
'End Function

'Public Sub RequeryMenifestCountIssued()
'On Error GoTo ErrD
'Dim tmpRec As New ADODB.Recordset
'Dim sqlStr As String
'    sqlStr = "SELECT * FROM vw_MenifestCountByJobHeader WHERE JobID='" & id_ & "'"
'    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    If tmpRec.EOF = False Then
'        menifestCountIssued_ = tmpRec("MenifestCount")
'    End If
'    tmpRec.Close
'    Set tmpRec = Nothing
'    Exit Sub
'ErrD:
'    GetMsg "", vbCritical
'    Exit Sub
'End Sub

'Private Function GetJobDetByWasteType(selType As Integer, carNo As Integer, selCol As Collection, genNoID As String) As Collection
'Dim wType As String
'Dim x%, tmpCol As Collection
'Dim aa As BWGJobDataDetailBox
'
'    wType = WasteTypes(selType).WasteType
'    Set tmpCol = New Collection
'    For x = 1 To colJobDet_.Count
'        If colJobDet_(x).WasteType = wType And colJobDet_(x).WasteGenNo = genNoID And colJobDet_(x).carNo = carNo And isExist(selCol, "'" & colJobDet_(x).ID & "'") = False Then
'            tmpCol.Add colJobDet_(x), "'" & colJobDet_(x).ID & "'"
'        End If
'    Next
'    Set GetJobDetByWasteType = tmpCol
'    Set tmpCol = Nothing
'End Function
