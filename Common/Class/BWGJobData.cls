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
Dim WorkDate_ As String
Dim workTime_ As String
Dim issuedStaffName_ As String
Dim issuedStaffID_ As String
Dim SaleStaffName_ As String
Dim Remdesc_ As String
Dim IssuedDate_ As String
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
Dim isCase_ As String
Dim isCopy_ As String
Dim CancelStaff_ As String
Dim cancelTime_ As String
Dim isNotActive_ As String
Dim PONo_ As String
Dim Numrow_ As String
Dim FromWork_ As String
Dim OrderType_ As String
Dim RefME_ As String
Dim IsChk_ As String
Dim CallOrd_ As String
Dim Callername_ As String
Dim CallerWay_ As String
Dim CancelRemark_ As String
Dim CovidWaste_ As String
Dim SaleDistrictIDOrder_ As String
Dim ServiceOrderID_ As String
Dim ServicePhoto_ As String
Dim AKPBoxCode_ As String
Dim AKPBoxCode2_ As String
Dim ProjectID_ As String
Dim PhotoS_ As String
Dim ComBme_ As String
Dim TruckType_  As String
Dim CostTrans_  As Double
Dim UserChk_ As String
Dim JobRef_ As String
Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Public Property Get id() As String
    id = id_
End Property
Public Property Let id(ByVal newval As String)
    id_ = Trim(newval)
End Property
Public Property Get ProjectID() As String
    ProjectID = ProjectID_
End Property
Public Property Let ProjectID(ByVal newval As String)
    ProjectID_ = Trim(newval)
End Property

Public Property Get JobRef() As String
    JobRef = JobRef_
End Property
Public Property Let JobRef(ByVal newval As String)
    If Trim(JobRef_) <> Trim(newval) Then
        JobRef_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get JobNo() As String
    JobNo = jobNo_
End Property
Public Property Let JobNo(ByVal newval As String)
    If Trim(jobNo_) <> Trim(newval) Then
        jobNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get ServiceOrderID() As String
    ServiceOrderID = ServiceOrderID_
End Property
Public Property Let ServiceOrderID(ByVal newval As String)
    If Trim(ServiceOrderID_) <> Trim(newval) Then
        ServiceOrderID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get ServicePhoto() As String
    ServicePhoto = ServicePhoto_
End Property
Public Property Let ServicePhoto(ByVal newval As String)
    If Trim(ServicePhoto_) <> Trim(newval) Then
        ServicePhoto_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CancelRemark() As String
    CancelRemark = CancelRemark_
End Property
Public Property Let CancelRemark(ByVal newval As String)
    If Trim(CancelRemark_) <> Trim(newval) Then
        CancelRemark_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CovidWaste() As String
    CovidWaste = CovidWaste_
End Property
Public Property Let CovidWaste(ByVal newval As String)
    If Trim(CovidWaste_) <> Trim(newval) Then
        CovidWaste_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SaleDistrictIDOrder() As String
    SaleDistrictIDOrder = SaleDistrictIDOrder_
End Property
Public Property Let SaleDistrictIDOrder(ByVal newval As String)
    If Trim(SaleDistrictIDOrder_) <> Trim(newval) Then
        SaleDistrictIDOrder_ = Trim(newval)
        isChanged_ = True
    End If
End Property


Public Property Get CustomerID() As String
    CustomerID = CustomerID_
End Property
Public Property Let CustomerID(ByVal newval As String)
    If Trim(CustomerID_) <> Trim(newval) Then
        CustomerID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerName() As String
    CustomerName = CustomerName_
End Property
Public Property Let CustomerName(ByVal newval As String)
    If Trim(CustomerName_) <> Trim(newval) Then
        CustomerName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteID() As String
    CustomerSiteID = CustomerSiteID_
End Property
Public Property Let CustomerSiteID(ByVal newval As String)
    If Trim(CustomerSiteID_) <> Trim(newval) Then
        CustomerSiteID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteName() As String
    CustomerSiteName = CustomerSiteName_
End Property
Public Property Let CustomerSiteName(ByVal newval As String)
    If Trim(CustomerSiteName_) <> Trim(newval) Then
        CustomerSiteName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteTel() As String
    CustomerSiteTel = customerSiteTel_
End Property
Public Property Let CustomerSiteTel(ByVal newval As String)
    If Trim(customerSiteTel_) <> Trim(newval) Then
        customerSiteTel_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteFax() As String
    CustomerSiteFax = customerSiteFax_
End Property
Public Property Let CustomerSiteFax(ByVal newval As String)
    If Trim(customerSiteFax_) <> Trim(newval) Then
        customerSiteFax_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteStaff() As String
    CustomerSiteStaff = customerSiteStaff_
End Property
Public Property Let CustomerSiteStaff(ByVal newval As String)
    If Trim(customerSiteStaff_) <> Trim(newval) Then
        customerSiteStaff_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteAumphurID() As String
    CustomerSiteAumphurID = customerSiteAumphurID_
End Property
Public Property Let CustomerSiteAumphurID(ByVal newval As String)
    If Trim(customerSiteAumphurID_) <> Trim(newval) Then
        customerSiteAumphurID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteProvinceID() As String
    CustomerSiteProvinceID = customerSiteProvinceID_
End Property
Public Property Let CustomerSiteProvinceID(ByVal newval As String)
    If Trim(customerSiteProvinceID_) <> Trim(newval) Then
        customerSiteProvinceID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerSiteProvinceName() As String
    CustomerSiteProvinceName = customerSiteProvinceName_
End Property
Public Property Let CustomerSiteProvinceName(ByVal newval As String)
    If Trim(customerSiteProvinceName_) <> Trim(newval) Then
        customerSiteProvinceName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WorkDate() As String
    WorkDate = WorkDate_
End Property
Public Property Let WorkDate(ByVal newval As String)
    If Trim(WorkDate_) <> Trim(newval) Then
        WorkDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WorkTime() As String
    WorkTime = workTime_
End Property
Public Property Let WorkTime(ByVal newval As String)
    If Trim(workTime_) <> Trim(newval) Then
        workTime_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedStaffName() As String
    IssuedStaffName = issuedStaffName_
End Property
Public Property Let IssuedStaffName(ByVal newval As String)
    If Trim(issuedStaffName_) <> Trim(newval) Then
        issuedStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CallOrd() As String
    CallOrd = CallOrd_
End Property
Public Property Let CallOrd(ByVal newval As String)
    If Trim(CallOrd_) <> Trim(newval) Then
        CallOrd_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CallerName() As String
    CallerName = Callername_
End Property
Public Property Let CallerName(ByVal newval As String)
    If Trim(Callername_) <> Trim(newval) Then
        Callername_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CallerWay() As String
    CallerWay = CallerWay_
End Property
Public Property Let CallerWay(ByVal newval As String)
    If Trim(CallerWay_) <> Trim(newval) Then
        CallerWay_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedStaffID() As String
    IssuedStaffID = issuedStaffID_
End Property
Public Property Let IssuedStaffID(ByVal newval As String)
    If Trim(issuedStaffID_) <> Trim(newval) Then
        issuedStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SaleStaffName() As String
    SaleStaffName = SaleStaffName_
End Property
Public Property Let SaleStaffName(ByVal newval As String)
    If Trim(SaleStaffName_) <> Trim(newval) Then
        SaleStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Remdesc() As String
    Remdesc = Remdesc_
End Property
Public Property Let Remdesc(ByVal newval As String)
    If Trim(Remdesc_) <> Trim(newval) Then
        Remdesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedDate() As String
    IssuedDate = IssuedDate_
End Property
Public Property Let IssuedDate(ByVal newval As String)
    If Trim(IssuedDate_) <> Trim(newval) Then
        IssuedDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedTime() As String
    IssuedTime = IssuedTime_
End Property
Public Property Let IssuedTime(ByVal newval As String)
    If Trim(IssuedTime_) <> Trim(newval) Then
        IssuedTime_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyID() As String
    CompanyID = companyID_
End Property
Public Property Let CompanyID(ByVal newval As String)
    If Trim(companyID_) <> Trim(newval) Then
        companyID_ = Trim(newval)
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
Public Property Let isCancel(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
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
Public Property Let WasteCreateNo(ByVal newval As String)
    If Trim(wasteCreateNo_) <> Trim(newval) Then
        wasteCreateNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteType() As String
    WasteType = wasteType_
End Property
Public Property Let WasteType(ByVal newval As String)
    If Trim(wasteType_) <> Trim(newval) Then
        wasteType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get MenifestCountAll() As Integer
    MenifestCountAll = menifestCountAll_
End Property
Public Property Let MenifestCountAll(ByVal newval As Integer)
    If menifestCountAll_ <> newval Then
        menifestCountAll_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get MenifestCountIssued() As Integer
    MenifestCountIssued = menifestCountIssued_
End Property
Public Property Let MenifestCountIssued(ByVal newval As Integer)
    If menifestCountIssued_ <> newval Then
        menifestCountIssued_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get SaleDistrictID() As String
    SaleDistrictID = saleDistrictID_
End Property
Public Property Let SaleDistrictID(ByVal newval As String)
    If Trim(saleDistrictID_) <> newval Then
        saleDistrictID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get UnderStaffID() As String
    UnderStaffID = underStaffID_
End Property
Public Property Let UnderStaffID(ByVal newval As String)
    If Trim(underStaffID_) <> newval Then
        underStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get UnderStaffName() As String
    UnderStaffName = underStaffName_
End Property
Public Property Let UnderStaffName(ByVal newval As String)
    If Trim(underStaffName_) <> newval Then
        underStaffName_ = Trim(newval)
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
Public Property Let isCompleted(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
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
Public Property Set JobDetails(ByVal newval As Collection)
    Set colJobDet_ = newval
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
Public Property Set Cars(ByVal newval As Collection)
    Set colJobCar_ = newval
    isChanged_ = True
End Property

Public Sub RequeryCar()
    Set colJobCar_ = Nothing
End Sub

Public Property Get isJob() As String
    isJob = isJob_
End Property
Public Property Let isJob(ByVal newval As String)
    If Trim(isJob_) <> newval Then
        isJob_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isCase() As String
    isCase = isCase_
End Property
Public Property Let isCase(ByVal newval As String)
    If isCase_ <> newval Then
        isCase_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get PhotoS() As String
    PhotoS = PhotoS_
End Property
Public Property Let PhotoS(ByVal newval As String)
    If PhotoS_ <> newval Then
        PhotoS_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get ComBme() As String
    ComBme = ComBme_
End Property
Public Property Let ComBme(ByVal newval As String)
    If ComBme_ <> newval Then
        ComBme_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get TruckType() As String
    TruckType = TruckType_
End Property
Public Property Let TruckType(ByVal newval As String)
    If TruckType_ <> newval Then
        TruckType_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get CostTrans() As String
    CostTrans = CostTrans_
End Property
Public Property Let CostTrans(ByVal newval As String)
    If CostTrans_ <> newval Then
        CostTrans_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get isCopy() As String
    isCopy = isCopy_
End Property
Public Property Let isCopy(ByVal newval As String)
    If Trim(isCopy_) <> newval Then
        isCopy_ = Trim(newval)
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Property Get CancelStaff() As String
    CancelStaff = CancelStaff_
End Property
Public Property Let CancelStaff(ByVal newval As String)
            CancelStaff_ = newval
End Property

Public Property Get cancelTime() As String
    cancelTime = cancelTime_
End Property
Public Property Let cancelTime(ByVal newval As String)
    cancelTime_ = newval
End Property
Public Property Get isNotActive() As String
    isNotActive = isNotActive_
End Property
Public Property Let isNotActive(ByVal newval As String)
    isNotActive_ = newval
End Property
Public Property Get PONo() As String
    PONo = PONo_
End Property
Public Property Let PONo(ByVal newval As String)
    PONo_ = newval
End Property
Public Property Get Numrow() As String
    Numrow = Numrow_
End Property
Public Property Let Numrow(ByVal newval As String)
    Numrow_ = newval
End Property
Public Property Get FromWork() As String
    FromWork = FromWork_
End Property
Public Property Let FromWork(ByVal newval As String)
    If Trim(FromWork_) <> Trim(newval) Then
            FromWork_ = Trim(newval)
            isChanged_ = True
    End If
End Property
Public Property Get OrderType() As String
    OrderType = OrderType_
End Property
Public Property Let OrderType(ByVal newval As String)
    If Trim(OrderType_) <> Trim(newval) Then
        OrderType_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get RefME() As String
    RefME = RefME_
End Property
Public Property Let RefME(ByVal newval As String)
    If Trim(RefME_) <> Trim(newval) Then
        RefME_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get IsChk() As String
    IsChk = IsChk_
End Property
Public Property Let IsChk(ByVal newval As String)
    If Trim(IsChk_) <> Trim(newval) Then
        IsChk_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get UserChk() As String
    UserChk = UserChk_
End Property
Public Property Let UserChk(ByVal newval As String)
    If Trim(UserChk_) <> Trim(newval) Then
        UserChk_ = Trim(newval)
        isChanged_ = True
    End If
End Property



Public Property Get AKPBoxCode() As String
    AKPBoxCode = AKPBoxCode_
End Property
Public Property Let AKPBoxCode(ByVal newval As String)
    If Trim(AKPBoxCode_) <> Trim(newval) Then
        AKPBoxCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get AKPBoxCode2() As String
    AKPBoxCode2 = AKPBoxCode2_
End Property
Public Property Let AKPBoxCode2(ByVal newval As String)
    If Trim(AKPBoxCode2_) <> Trim(newval) Then
        AKPBoxCode2_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Sub Save(ByVal BoxCode As String, ByVal BoxCode2 As String)
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset, rs As ADODB.Recordset
Dim RSTruck As ADODB.Recordset, RSCar As ADODB.Recordset, RSBox As ADODB.Recordset
Dim RSBox2 As ADODB.Recordset
Dim RSstm As ADODB.Recordset
Dim RSBoxAmt As ADODB.Recordset
Dim x%, isNew As Boolean
Dim sqlStr As String
Dim sqlStrTruck As String
Dim SqlBox As String
Dim sqlBox2 As String
Dim sqlBoxAmt As String
Dim Sql As String
Dim CarUse As Integer
Dim CarUsed As Integer
Dim OldCarUsed As Integer

    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbJobDataHeader WHERE JobID='" & id_ & "'"
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
    If (isNew = False And Trim("" & tmpRec("WorkDate")) <> Trim(WorkDate_)) Or isNew = True Then
        tmpRec("WorkDate") = Trim(WorkDate_)
        'DBConnExc "Update tbJobDataTimeTable set WorkDate='" & workDate_ & "' WHERE JobDataCarID in (Select JobDataCarID From tbJobDataCarUsed WHERE JobID='" & id_ & "')"
    End If
    tmpRec("WorkTime") = Trim(workTime_)
    tmpRec("IssuedStaffID") = Trim(issuedStaffID_)
    tmpRec("UnderSaleStaffID") = Trim(underStaffID_)
    tmpRec("SaleStaffName") = Trim(SaleStaffName_)
    tmpRec("RemDesc") = Trim(Remdesc_)
    tmpRec("IssuedDate") = Trim(IssuedDate_)
    tmpRec("IssuedTime") = Trim(IssuedTime_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("isCancel") = Trim(isCancel_)
    tmpRec("WasteCreateNo") = Trim(wasteCreateNo_)
    tmpRec("WasteType") = Trim(wasteType_)
   If Trim(isCompleted_) = "" Then isCompleted_ = "N"
    tmpRec("isCompleted") = Trim(isCompleted_)
    tmpRec("PONo") = Trim(PONo_)
    tmpRec!RefME = Trim(RefME_)
    tmpRec!CallOrd = Trim(CallOrd_)
    tmpRec("Callername") = Trim(Callername_)
    tmpRec("ProjectID") = Trim(ProjectID_)
    tmpRec!CallerWay = Trim(CallerWay_)
    'tmpRec!IsChk = Trim(IsChk_)
    tmpRec!SaleDistrictIDOrder = Trim(SaleDistrictIDOrder_)
    tmpRec!CancelRemark = Trim(CancelRemark_)
    tmpRec!CovidWaste = Trim(CovidWaste_)
    tmpRec!JobRef = Trim(JobRef_)
    
    
    tmpRec!isJobs = IIf(isJob_ = "", "N", isJob_)
    tmpRec!WasteCaseID = isCase_
    tmpRec!FromWork = Trim(FromWork_)
    tmpRec!OrderType = IIf(OrderType_ = "", "N", OrderType_)
    tmpRec!Service_OrderID = Trim(ServiceOrderID_)
    If Trim(BoxCode) <> "" Then
    tmpRec!AKPBoxCode = Trim(BoxCode)
    End If
    If Trim(BoxCode2) <> "" Then
    tmpRec!AKPBoxCode2 = Trim(BoxCode2)
    End If
    tmpRec!LastUserUpdate = Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName)
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    tmpRec!PhotoS = IIf(PhotoS = "", "N", PhotoS)
    tmpRec!ComBme = IIf(ComBme = "", "N", ComBme)
    tmpRec!TruckType = TruckType
    tmpRec!CostTrans = CostTrans
    
    
    
    tmpRec.Update
    tmpRec.Requery
    jobNo_ = Trim("" & tmpRec("JobNo"))
       
    DBConn.Execute "UPDATE tbJobDataHeader SET LastUpdate = CONVERT(DATETIME, GETDATE(), 103) WHERE JobID='" & id_ & "'"
    DBConn.Execute "UPDATE tbJobProject SET ProjectStatus = 'P' WHERE ProjectID='" & ProjectID_ & "'"
    
    'เช็คว่ามีการ Confirm รึยัง
    If ServiceOrderID_ <> "" Then
    Dim sqlstm As String
        sqlstm = "SELECT * FROM Service_OrderHD WHERE orderid='" & ServiceOrderID_ & "' and confirmorder='Y' "
        Set RSstm = New ADODB.Recordset
        GetDBConnection
        RSstm.open sqlstm, DBConn, adOpenDynamic, adLockOptimistic
        Dim Lnk As String
        Lnk = "http://www.akpcustomer.com/Order/ConfrimOrderForCustomer?id=" & ServiceOrderID_ & ""
        Dim r As Long
        
        '--------ถ้ายังไม่ confirm ทำการอัพเดทการ Confirm แล้วส่งเมลล์แจ้ง
        If RSstm.RecordCount <= 0 Then
            DBConn.Execute "UPDATE Service_OrderHD SET confirmorder = 'Y' WHERE orderid='" & ServiceOrderID_ & "'"
            r = ShellExecute(0, "open", Lnk, 0, 0, 1)
        End If
    End If

'    tmpRec.Close

        If Not colJobDet_ Is Nothing Then
            DBConnExc "Update tbJobDataDetail Set MarkForDelete='Y' WHERE JobID='" & id_ & "'"
            For x = 1 To colJobDet_.Count
                colJobDet_(x).RowID = x
                colJobDet_(x).Save
            Next
            DBConnExc "Delete tbJobDataDetail WHERE MarkForDelete='Y' And JobID='" & id_ & "'"
        End If
        If Not colJobCar_ Is Nothing Then
            DBConnExc "Update tbJobDataCarUsed Set MarkForDelete='Y' WHERE JobID='" & id_ & "'"
            For x = 1 To colJobCar_.Count
                colJobCar_(x).Save Trim(WorkDate_), isCopy_
            Next
            DBConnExc "Delete tbJobDataTimeTable WHERE JobDataCarID in (Select JobDataCarID From tbJobDataCarUsed WHERE MarkForDelete='Y' And JobID='" & id_ & "')"
            DBConnExc "Delete tbJobDataCarUsed WHERE MarkForDelete='Y' And JobID='" & id_ & "'"
        End If
        
    'ลบจำนวนรถในหน้ารอจัดรถ
    Dim TruckIDs As String
    If isNew = True Then
        sqlStrTruck = "SELECT TruckSubTypeID FROM tbJobDataCarUsed WHERE JobID='" & id_ & "'"
        Set RSTruck = New ADODB.Recordset
        GetDBConnection
        RSTruck.open sqlStrTruck, DBConn, adOpenDynamic, adLockOptimistic
        
        TruckIDs = RSTruck!TruckSubTypeID
        
            If RSTruck.RecordCount > 0 Then
                Sql = "SELECT * FROM tbCarDaily WHERE subtypeid='" & RSTruck!TruckSubTypeID & "' and workdate='" & Trim(WorkDate_) & "' "
                Set RSCar = New ADODB.Recordset
                GetDBConnection
                RSCar.open Sql, DBConn, adOpenDynamic, adLockOptimistic
            If RSCar.RecordCount > 0 Then
                If TruckIDs = "33D46FC64F3D450897D1F72986EAE700" Then
                    If RSCar!CarUse <= 0 Then
                        Dim sql2 As String
                        Dim RS2 As ADODB.Recordset
                            sql2 = "SELECT * FROM tbCarDaily WHERE subtypeid='058A0646AD8D401BB7884A63C7F1F92E' and workdate='" & Trim(WorkDate_) & "' "
                            Set RS2 = New ADODB.Recordset
                            GetDBConnection
                            RS2.open sql2, DBConn, adOpenDynamic, adLockOptimistic
                            If RS2.RecordCount > 0 Then
                                If IsNull(RS2!CarUsed) Then
                                    OldCarUsed = 0
                                Else
                                    OldCarUsed = RS2!CarUsed
                                End If
                                CarUse = CDbl(RS2!CarUse) - 1
                                CarUsed = CDbl(OldCarUsed) + 1
                                DBConnExc "Update tbCarDaily set CarUse='" & CarUse & "', CarUsed='" & CarUsed & "'  WHERE CarDailyID='" & RS2!CarDailyID & "' "
                            End If
                    Else
                        If IsNull(RSCar!CarUsed) Then
                            OldCarUsed = 0
                        Else
                            OldCarUsed = RSCar!CarUsed
                        End If
                    CarUse = CDbl(RSCar!CarUse) - 1
                    CarUsed = CDbl(OldCarUsed) + 1
                    DBConnExc "Update tbCarDaily set CarUse='" & CarUse & "', CarUsed='" & CarUsed & "'  WHERE CarDailyID='" & RSCar!CarDailyID & "' "
                End If
            Else
                    CarUse = CDbl(RSCar!CarUse) - 1
                    CarUsed = CDbl(OldCarUsed) + 1
                    DBConnExc "Update tbCarDaily set CarUse='" & CarUse & "', CarUsed='" & CarUsed & "'  WHERE CarDailyID='" & RSCar!CarDailyID & "' "
            End If
        End If
    End If
    End If

'ลบจำนวน BOX ในหน้ารอจัดรถ
'    If isNew = True Then
    If BoxCode = "" Then BoxCode = "0"
    If BoxCode2 = "" Then BoxCode2 = "0"
        Dim Codexx As String
        SqlBox = "SELECT * FROM vw_BoxAKP WHERE BoxCode = '" & BoxCode & "' and WorkDate='" & Trim(WorkDate_) & "' "
        Set RSBox = New ADODB.Recordset
        GetDBConnection
        RSBox.open SqlBox, DBConn, adOpenDynamic, adLockOptimistic
            If RSBox.RecordCount > 0 Then
            Dim OldBoxUse As Integer
            Dim UseBox As Integer
                    If IsNull(RSBox!boxUse) Then
                        OldBoxUse = 0
                    Else
                        OldBoxUse = RSBox!boxUse
                    End If
                    UseBox = CDbl(OldBoxUse) - 1
                    DBConnExc "Update tbAKPBoxUse set BoxUse='" & UseBox & "' WHERE ID='" & RSBox!useID & "' "
            End If
        sqlBox2 = "SELECT * FROM vw_BoxAKP WHERE BoxCode = '" & BoxCode2 & "' and WorkDate='" & Trim(WorkDate_) & "' "
        Set RSBox2 = New ADODB.Recordset
        GetDBConnection
        RSBox2.open sqlBox2, DBConn, adOpenDynamic, adLockOptimistic
            If RSBox2.RecordCount > 0 Then
            Dim OldBoxUse2 As Integer
            Dim UseBox2 As Integer
                    If IsNull(RSBox2!boxUse) Then
                        OldBoxUse2 = 0
                    Else
                        OldBoxUse2 = RSBox2!boxUse
                    End If
                    UseBox2 = CDbl(OldBoxUse2) - 1
                    DBConnExc "Update tbAKPBoxUse set BoxUse='" & UseBox2 & "' WHERE ID='" & RSBox2!useID & "' "
            End If
                
            'คืนค่า Box ตัวเก่ากรณีที่เปลี่ยน Box
                Dim UseboxAmt As Integer
                If Trim(BoxCode) <> "" Then
                    If BoxCode <> frmJobOpen.Obc1.Caption Then
                        UseboxAmt = 0
                        Codexx = Trim(frmJobOpen.Obc1.Caption)
                        If Codexx = "" Then Codexx = 0
                        sqlBoxAmt = "SELECT * FROM vw_BoxAKP WHERE BoxCode = '" & Codexx & "' and WorkDate='" & Trim(WorkDate_) & "' "
                        Set RSBoxAmt = New ADODB.Recordset
                        GetDBConnection
                        RSBoxAmt.open sqlBoxAmt, DBConn, adOpenDynamic, adLockOptimistic
                            If RSBoxAmt.RecordCount > 0 Then
                                UseboxAmt = RSBoxAmt!boxUse + 1
                                DBConnExc "Update tbAKPBoxUse set BoxUse='" & UseboxAmt & "' WHERE ID='" & RSBoxAmt!useID & "' "
                            End If
                    End If
                End If
                If Trim(BoxCode2) <> "" Then
                    If BoxCode2 <> frmJobOpen.Obc2.Caption Then
                    Dim sqlx As String
                    Dim RSxx As ADODB.Recordset
                        UseboxAmt = 0
                        Codexx = Trim(frmJobOpen.Obc2.Caption)
                        If Codexx = "" Then Codexx = 0
                        sqlx = "SELECT * FROM vw_BoxAKP WHERE BoxCode = '" & Codexx & "' and WorkDate='" & Trim(WorkDate_) & "' "
                        Set RSxx = New ADODB.Recordset
                        GetDBConnection
                        RSxx.open sqlx, DBConn, adOpenDynamic, adLockOptimistic
                            If RSxx.RecordCount > 0 Then
                                UseboxAmt = RSxx!boxUse + 1
                                DBConnExc "Update tbAKPBoxUse set BoxUse='" & UseboxAmt & "' WHERE ID='" & RSxx!useID & "' "
                            End If
                    End If
                End If
    

    Set tmpRec = Nothing
    isChanged_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub
Public Function SaveUserChk() As String
On Error GoTo delErr:
    DBConnExc "Update  tbJobDataHeader set  UserChk='" & UserChk_ & "' ,IsChk='" & IsChk_ & "'  WHERE JobID='" & id_ & "'"
    SaveUserChk = ""
    Exit Function
delErr:
    SaveUserChk = err.Description
End Function
Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbJobDataHeader WHERE JobID='" & id_ & "'"
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
    WorkDate_ = Trim("" & selRec("WorkDate"))
    workTime_ = Trim("" & selRec("WorkTime"))
    issuedStaffID_ = Trim("" & selRec("IssuedStaffID"))
    issuedStaffName_ = Trim("" & selRec("IssuedStaffName"))
    SaleStaffName_ = Trim("" & selRec("SaleStaffName"))
    Remdesc_ = Trim("" & selRec("RemDesc"))
    IssuedDate_ = Trim("" & selRec("IssuedDate"))
    IssuedTime_ = Trim("" & selRec("IssuedTime"))
    companyID_ = Trim("" & selRec("CompanyID"))
    isCancel_ = Trim("" & selRec("isCancel"))
    If isExist(selRec, "SaleDistrictID") Then saleDistrictID_ = Trim("" & selRec("SaleDistrictID"))
    underStaffID_ = Trim("" & selRec("UnderSaleStaffID"))
    underStaffName_ = Trim("" & selRec("StaffFName")) & " " & Trim("" & selRec("StaffLName"))
    If IsNull(selRec("MenifestCount")) = False Then menifestCountIssued_ = Trim("" & selRec("MenifestCount"))
    If IsNull(selRec("MenifestCountAll")) = False Then menifestCountAll_ = Trim("" & selRec("MenifestCountAll"))
    wasteCreateNo_ = Trim("" & selRec("WasteCreateNo"))
    wasteType_ = Trim("" & selRec("WasteType"))
    isCompleted_ = Trim("" & selRec("isCompleted"))
    PONo_ = Trim("" & selRec("PONo"))
    Numrow_ = Trim("" & selRec("NumberRow"))
    isJob_ = selRec!isJobs
    If IsNull(selRec("WasteCaseID")) = False Then isCase_ = Trim("" & selRec("WasteCaseID"))
    isNotActive_ = selRec!isNotActive
    FromWork_ = Trim("" & selRec("FromWork"))
    OrderType_ = Trim("" & selRec("OrderType"))
    PhotoS_ = Trim("" & selRec("PhotoS"))
    ComBme_ = Trim("" & selRec("ComBme"))
    RefME_ = Trim("" & selRec("RefME"))
    IsChk_ = Trim("" & selRec("IsChk"))
    UserChk_ = Trim("" & selRec("UserChk"))
    CallOrd_ = Trim("" & selRec("CallOrd"))
    Callername_ = Trim("" & selRec("Callername"))
    CallerWay_ = Trim("" & selRec("CallerWay"))
    CancelRemark_ = Trim("" & selRec("CancelRemark"))
    CovidWaste_ = Trim("" & selRec("CovidWaste"))
    SaleDistrictIDOrder_ = Trim("" & selRec("SaleDistrictIDOrder"))
    ServiceOrderID_ = Trim("" & selRec("Service_OrderID"))
    AKPBoxCode_ = Trim("" & selRec("AKPBoxCode"))
    AKPBoxCode2_ = Trim("" & selRec("AKPBoxCode2"))
    TruckType_ = Trim("" & selRec("TruckType"))
    If IsNull(selRec("CostTrans")) = False Then CostTrans_ = Trim("" & selRec("CostTrans"))
    If isExist(selRec, "ProjectID") Then ProjectID_ = IIf(IsNull(selRec!ProjectID), "", selRec!ProjectID)
End Sub

Public Sub CancelJob()
    Dim tSql As String
    DBConnExc "Update tbJobDataHeader Set isCancel='Y', cancelstaff = '" & CancelStaff_ & "',CancelRemark ='" & CancelRemark_ & "', cancelTime='" & cancelTime_ & "'  WHERE JobID='" & id_ & "'"
    isCancel_ = "Y"
    GetDBConnection
    
    DBConnExc "Update Service_OrderHD set isCancel='Y', CancelDateTime = GETDATE() WHERE OrderID='" & ServiceOrderID_ & "'"
    GetDBConnection
End Sub

Public Sub CloseJob()
    DBConnExc "Update tbJobDataHeader Set isCompleted='Y' WHERE JobID='" & id_ & "'"
    isCompleted_ = "Y"
End Sub

Public Function DuplicateJob(newWorkDate As String) As BWGJobData
Dim newJob As New BWGJobData
Dim x%, newJobDet As BWGJobDataDetail, y%
Dim newJobCar As BWGJobDataCarUsed
Dim newTimeTable As BWGJobDataTimeTable
Dim tmpS As New BWGSearchManager
Dim colCustomer As Collection, JobCustomer As New BWGCustomer
Dim selDate As String
    If Trim(newWorkDate) = "" Then
        selDate = WorkDate_
    Else
        selDate = newWorkDate
    End If
    With newJob
        .JobRef = id_
        .id = GetGUID
        .JobNo = jobNo_
        .CustomerID = CustomerID_
        .CustomerName = CustomerName_
        .CustomerSiteID = CustomerSiteID_
        .CustomerSiteName = CustomerSiteName_
        .CustomerSiteTel = customerSiteTel_
        .CustomerSiteFax = customerSiteFax_
        .CustomerSiteStaff = customerSiteStaff_
        .CustomerSiteAumphurID = customerSiteAumphurID_
        Set colCustomer = tmpS.CustomerSearch("CustomerID = '" & CustomerID_ & "'")
        Set JobCustomer = colCustomer(1)
        .UnderStaffID = JobCustomer.UnderSaleStaffID
        .UnderStaffName = JobCustomer.UnderSaleName
        Set JobCustomer = Nothing
        Set colCustomer = Nothing
        .WorkDate = selDate
        .WorkTime = workTime_
        .WasteCreateNo = wasteCreateNo_
        .IssuedStaffID = issuedStaffID_
        .IssuedStaffName = issuedStaffName_
        .SaleStaffName = SaleStaffName_
        .Remdesc = Remdesc_
        .IssuedDate = IssuedDate_
        .IssuedTime = IssuedTime_
        .CompanyID = companyID_
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
                    newJobDet.id = GetGUID
                    newJobDet.JobID = newJob.id
                    newJobDet.isPriceIncTrans = .isPriceIncTrans
                    newJobDet.JobNo = ""
                    newJobDet.Remdesc = .Remdesc
                    newJobDet.RowID = .RowID
                    newJobDet.wastedataID = .wastedataID
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
                    ''******************** แก้ไข ข้อมูลราคาและหน่วยไม่ถูกก็อปปี้ ทำให้ไม่คำนวณข้อมูลในการออก Invoice
                    newJobDet.JobTreatmentRate = .JobTreatmentRate
                    newJobDet.JobTreatmentUnitID = .JobTreatmentUnitID
                    newJobDet.JobTransportFee = .JobTransportFee
                    newJobDet.JobTransUnitID = .JobTransUnitID
                    ''**************************************************************
                End With
                .JobDetails.Add newJobDet, "'" & newJobDet.id & "'"
                Set newJobDet = Nothing
            Next
        End If
        If colJobCar_ Is Nothing Then Set colJobCar_ = tmpS.JobDataCarSearch("JobID='" & id_ & "'")
        If Not colJobCar_ Is Nothing Then
            For x = 1 To colJobCar_.Count
                Set newJobCar = New BWGJobDataCarUsed
                With colJobCar_(x)
                
                    Dim TruckSubTypeID2 As String
                    TruckSubTypeID2 = ""
                    If .TruckSubTypeID = "33D46FC64F3D450897D1F72986EAE700" Then TruckSubTypeID2 = "F0D9A37CFC034EB89BF0F0DBA15D9523"
                    If .TruckSubTypeID = "F0D9A37CFC034EB89BF0F0DBA15D9523" Then TruckSubTypeID2 = "33D46FC64F3D450897D1F72986EAE700"
                    If .TruckSubTypeID = "2EFEDFB372C448E3B93C786DA8613F6A" Then TruckSubTypeID2 = "CD50C04FD207477BBA66216557868208"
                    If .TruckSubTypeID = "CD50C04FD207477BBA66216557868208" Then TruckSubTypeID2 = "2EFEDFB372C448E3B93C786DA8613F6A"
                
                    TruckAvailable.TruckAvailable = 0
                    TruckAvailable.TruckTotal = 0
                    TruckAvailableCount newJob.WorkDate, .TruckSubTypeID, TruckSubTypeID2
                    If TruckAvailable.TruckTotal > 0 Then
                        If TruckAvailable.TruckAvailable < .TruckCount Then
                            MsgBox "จำนวนรถ """ & newJobCar.TruckSubTypeName & """ ไม่พอ ไม่สามารถลงจำนวนรถได้" & vbCrLf & "กรุณาตรวจสอบจำนวนรถของวันนี้กับฝ่ายขนส่งอีกครั้ง, จำนวนรถคงเหลือ (" & TruckAvailable.TruckAvailable & ")", vbInformation
                            Exit For
                        End If
                    End If
                    
                    newJobCar.DestinationDesc = .DestinationDesc
                    newJobCar.id = GetGUID
                    newJobCar.JobID = newJob.id
                    newJobCar.Remark = .Remark
                    newJobCar.RowID = .RowID
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
                        newTimeTable.JobDataCarID = newJobCar.id
                        newTimeTable.StartTime = .TimeTables(y).StartTime
                        newTimeTable.TripNo = .TimeTables(y).TripNo
                        newTimeTable.WorkDate = selDate
                        newTimeTable.WorkerCount = .WorkerPerCar
                        newJobCar.TimeTables.Add newTimeTable, "'" & newTimeTable.JobDataCarID & "_" & newTimeTable.TripNo & "'"
                        Set newTimeTable = Nothing
                    Next
                End With
                .Cars.Add newJobCar, "'" & newJobCar.id & "'"
               
                Set newJobCar = Nothing
            Next
        End If
    End With
    Set DuplicateJob = newJob
    Set newJob = Nothing
    Set tmpS = Nothing
End Function

Public Function GenerateMenifest() As Integer
Dim tmpCar As BWGJobDataCarUsed, tmpTimeTable As BWGJobDataTimeTable
Dim tmpMenifest As BWGMenifest, tmpCon As ADODB.Connection
Dim x%, y%, iCount%, z%, colWaste As Collection, aa%
Dim tmpCol As Collection, tmpS As New BWGSearchManager
Dim enumWType() As String, colGenNo As Collection, bb%
    Set colGenNo = tmpS.WasteGenNoSearch("JobID='" & id_ & "'")
    ReDim enumWType(WasteTypes.Count)
    For x = 1 To WasteTypes.Count
        enumWType(x) = WasteTypes(x).WasteType
    Next
            
    For x = 1 To colJobCar_.Count
        Set tmpCar = colJobCar_(x)
        For aa = 1 To WasteTypes.Count
            For y = 1 To tmpCar.TimeTables.Count
                Set tmpTimeTable = tmpCar.TimeTables(y)
                For bb = 1 To colGenNo.Count
                    Set colWaste = GetJobDetByWasteType(aa, tmpCar.carNo, tmpTimeTable.JobDetails, colGenNo(bb).WasteCreateNo)
                    If colWaste.Count > 0 Then
                        If Trim(tmpTimeTable.CarRegisNo) <> "" Then
                            'If Trim(tmpTimeTable.MenifestID) = "" Then
                            '    tmpTimeTable.MenifestID = GetGUID
                            '    DBConnExc "Update tbJobDataTimeTable Set MenifestID='" & tmpTimeTable.MenifestID & "' WHERE JobDataCarID='" & tmpTimeTable.JobDataCarID & "' And TripNo=" & tmpTimeTable.TripNo
                            'End If
                            Set tmpMenifest = New BWGMenifest
                            With tmpMenifest
                                .id = GetGUID
                                .CreatedFromJobID = id_
                                .CompanyID = companyID_
                                .CustomerID = CustomerID_
                                .CustSiteID = CustomerSiteID_
                                .WasteGenNo = colGenNo(bb).WasteCreateNo
                                .NonOrHz = enumWType(aa)
                                .TelNo = customerSiteTel_
                                .FaxNo = customerSiteFax_
                                .PONo = PONo_
                                
                                '*** ตรวจสอบการคิดน้ำหนักจากผู้รับกำจัด หรือผู้ก่อนกำเนิด
                                Dim tSql As String
                                Dim rs As New ADODB.Recordset
                                tSql = "SELECT * FROM tbCustomer_TR_ChargeWType WHERE CustomerID = '" & CustomerID_ & "'"
                                Set rs = New ADODB.Recordset
                                Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
                                If rs.RecordCount > 0 Then
                                    Select Case rs!ChargeWType
                                        Case "A"
                                            .ChargeWType = "A"
                                        Case "B"
                                            .ChargeWType = "B"
                                        Case Else
                                            .ChargeWType = "A"
                                    End Select
                                Else
                                    .ChargeWType = "A"
                                End If
                                '*************************
                                
                                .ErContact = customerSiteStaff_
                                If Trim(tmpCar.TruckTypeName) <> "รถโรงงานลูกค้า" Then
                                    .TransportName1 = "บริษัท  อัคคีปราการ  จำกัด (มหาชน)"
                                    .TransRegisNo1 = "DIW-T-085800068"
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
                                .TruckSubTypeID = tmpCar.TruckSubTypeID
                                .CarRegisID = tmpTimeTable.CarRegisNo
                                .OutWorkerCount = tmpCar.WorkerPerCar
                                .FromProvinceID = customerSiteProvinceID_
                                .FromProvinceName = customerSiteProvinceName_
                                .OpenStaffID = CurrentUser.id
                                .OpenDate = TodayDate
                                .WorkDate = WorkDate_
                                .isClosed = False
                                .IsCanceled = False
                                .JobDataCarID = tmpTimeTable.JobDataCarID
                                .TripNo = tmpTimeTable.TripNo
                                .DriverName = tmpTimeTable.DriverName
                                
                                
                                If Trim(.SaleDistricIDOrder) <> "" Then
                                    .SaleDistrictID = .SaleDistricIDOrder
                                Else
                                    .SaleDistrictID = saleDistrictID_
                                End If
                                
                                .UnderStaffID = underStaffID_
                                .UnderStaffName = underStaffName_
                                For z = 1 To colWaste.Count
                                    If colWaste(z).carNo = tmpCar.carNo And colJobDet_(z).menifestID = "" Then
                                        tmpTimeTable.JobDetails.Add colWaste(z), "'" & colWaste(z).id & "'"
                                        .JobDetails.Add colWaste(z), "'" & colWaste(z).id & "'"
                                        If Trim(.TSDF1ID) = "" Then
                                            Set tmpCol = tmpS.TSDFSearch("TSDFID='" & colWaste(z).TSDFID & "'")
                                            If tmpCol.Count > 0 Then
                                                .TSDFName1 = tmpCol(1).CompanyName
                                                .TSDF1ID = tmpCol(1).id
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
                                .isJob = isJob_
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
Dim sqlStr As String
    sqlStr = "SELECT * FROM vw_MenifestCountByJobHeader WHERE JobID='" & id_ & "'"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF = False Then
        menifestCountIssued_ = tmpRec("MenifestCount")
    End If
    tmpRec.Close
    Set tmpRec = Nothing
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Private Function GetJobDetByWasteType(selType As Integer, carNo As Integer, selCol As Collection, genNoID As String) As Collection
Dim wType As String
Dim x%, tmpCol As Collection
Dim aa As BWGJobDataDetail

    wType = WasteTypes(selType).WasteType
    Set tmpCol = New Collection
    For x = 1 To colJobDet_.Count
        If colJobDet_(x).WasteType = wType And colJobDet_(x).WasteGenNo = genNoID And colJobDet_(x).carNo = carNo And isExist(selCol, "'" & colJobDet_(x).id & "'") = False Then
            tmpCol.Add colJobDet_(x), "'" & colJobDet_(x).id & "'"
        End If
    Next
    Set GetJobDetByWasteType = tmpCol
    Set tmpCol = Nothing
End Function
