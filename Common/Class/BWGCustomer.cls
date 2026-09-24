VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim CustomerCode_ As String
Dim custName_ As String
Dim custNameCer_ As String
Dim custShortName_ As String
Dim CustAddr_ As String
Dim SubContact_ As String
Dim tumbolID_ As String
Dim tumbolName_ As String
Dim aumphurID_ As String
Dim aumphurName_ As String
Dim provinceID_ As String
Dim provinceName_ As String
Dim zipCode_ As String
Dim Country_ As String
Dim TelNo_ As String
Dim faxNo_ As String
Dim website_ As String
Dim emailAddr_ As String
Dim contactPerson_ As String
Dim positionName_ As String
Dim deptName_ As String
Dim industryTypeID_ As String
Dim industryTypeName_ As String
Dim industryTypeIDNew_ As String
Dim TypeDesc_ As String
Dim zoneID_ As String
Dim zoneName_ As String
Dim RegisterNo_ As String
Dim WasteManageID_ As String
Dim taxNo_ As String
Dim isAccountOK_ As String
Dim custType_ As String
Dim parentCustID_ As String
Dim parentCustCode_ As String
Dim parentCustName_ As String
Dim createdStaffID_ As String
Dim createdStaffName_ As String
Dim createdDate_ As String
Dim underSaleStaffID_ As String
Dim underSaleName_ As String
Dim underSaleTelNo_ As String
Dim custGroupID_ As String
Dim custGroupName_ As String
Dim colWasteItem_ As Collection
Dim colSiteAddr_ As Collection
Dim colBillAddr_ As Collection
Dim colContract_ As Collection
Dim colAuction_ As Collection
Dim colCustStaff_ As Collection
Dim colBillingStaff_ As Collection
Dim CmmQuo_ As Collection
Dim colOtherRemark_ As Collection
Dim billingInfo_ As BWGCustomerBillInfo
Dim saleDistrictID_ As String
Dim saleDistrictName_ As String
Dim saleZoneID_ As String
Dim saleZoneName_ As String
Dim SaleAP_ As String
Dim Sale2ID_ As String
Dim Sale2Name_ As String
Dim custRemark_ As String
Dim isChanged_ As Boolean
Dim Status_ As String
Dim Refcode_ As String
Dim BeginDate_ As String
Dim GPSRequest_ As String

Dim isNew_ As String
Dim Comment_ As String
Dim CommentDate_ As String
Dim cancelDate_ As String
Dim CancelStaff_ As String
Dim DTeam_ As String
Dim CommentTimeNo_ As String
Dim Creditterm_ As String
Dim CustdeptName_ As String

Dim LastUser_ As String
Dim LastUpdate_ As String
Dim LastUserTR_ As String
Dim LastUpdateTR_ As String
Dim ischeck_ As String
Dim isWeigth_ As String

Dim isBox_ As String

Dim isLUGGERBOX_ As String
Dim LUGBOX_ As String
Dim LUGBOXFrom_ As String
Dim LUGBOXBack_ As String
Dim LUGBOXBackV_ As String
Dim LUGBoxW_ As String

Dim isROLGERBOX_ As String
Dim ROLBOX_ As String
Dim ROLBOXFrom_ As String
Dim ROLBOXBack_ As String
Dim ROLBOXBackv_ As String
Dim ROLBoxW_ As String

Dim isOneTBOX_ As String
Dim OneT_ As String
Dim OneTFrom_ As String
Dim OneTBack_ As String
Dim OneTBackv_ As String
Dim OneTBoxW_ As String

Dim isTwoTBOX_ As String
Dim TwoT_ As String
Dim TwoTFrom_ As String
Dim TwoTBack_ As String
Dim TwoTBackv_ As String
Dim TwoTBoxW_ As String

Dim isTruckBOX_ As String
Dim TruckBOX_ As String
Dim TruckBOXFrom_ As String
Dim TruckBOXBack_ As String
Dim TruckBOXBackv_ As String
Dim TruckBOXW_ As String

Dim Lack_ As String
Dim LackFrom_ As String
Dim LackBack_ As String
Dim LackTo_ As String
Dim LackW_ As String

Dim BoutS_ As String
Dim BoutSFrom_ As String
Dim BoutSBack_ As String
Dim BoutSTo_ As String
Dim BoutSW_ As String

Dim WaterBout_ As String
Dim WaterBoutFrom_ As String
Dim WaterBoutBack_ As String
Dim WaterBoutTo_ As String
Dim WaterBoutW_ As String

Dim BoutCage_ As String
Dim BoutCageFrom_ As String
Dim BoutCageBack_ As String
Dim BoutCageTo_ As String
Dim BoutCageW_ As String

Dim isHaz_ As String
Dim isNon_ As String
Dim PriceBox_ As String
Dim WasteSK2Exp_ As String
Dim OverWeight_ As String
Dim BoxDesc_ As String
Dim StpTrn_ As String
Dim isVIPCR_ As String
Dim CRStaffID_ As String
Dim CRStaffName_ As String
Dim PicOfDay_ As Integer
Dim Digit14_ As String
Public Property Get Digit14() As String
    Digit14 = Digit14_
End Property
Public Property Let Digit14(ByVal newval As String)
    If Trim(Digit14_) <> Trim(newval) Then
        Digit14_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get PicOfDay() As Integer
    PicOfDay = PicOfDay_
End Property
Public Property Let PicOfDay(ByVal newval As Integer)
    If Trim(PicOfDay_) <> Trim(newval) Then
        PicOfDay_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CRStaffName() As String
    CRStaffName = CRStaffName_
End Property
Public Property Let CRStaffName(ByVal newval As String)
    If Trim(CRStaffName_) <> Trim(newval) Then
        CRStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CRStaffID() As String
    CRStaffID = CRStaffID_
End Property
Public Property Let CRStaffID(ByVal newval As String)
    If Trim(CRStaffID_) <> Trim(newval) Then
        CRStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property
Public Property Get WasteSK2Exp() As String
    WasteSK2Exp = WasteSK2Exp_
End Property
Public Property Let WasteSK2Exp(ByVal newval As String)
    If Trim(WasteSK2Exp_) <> Trim(newval) Then
        WasteSK2Exp_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Refcode() As String
    Refcode = Refcode_
End Property
Public Property Let Refcode(ByVal newval As String)
    If Trim(Refcode_) <> Trim(newval) Then
        Refcode_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get GPSRequest() As String
    GPSRequest = GPSRequest_
End Property
Public Property Let GPSRequest(ByVal newval As String)
    If Trim(GPSRequest_) <> Trim(newval) Then
        GPSRequest_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get OverWeight() As String
    OverWeight = OverWeight_
End Property
Public Property Let OverWeight(ByVal newval As String)
    If Trim(OverWeight_) <> Trim(newval) Then
        OverWeight_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BeginDate() As String
    BeginDate = BeginDate_
End Property
Public Property Let BeginDate(ByVal newval As String)
    If Trim(BeginDate_) <> Trim(newval) Then
        BeginDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CustomerCode() As String
    CustomerCode = CustomerCode_
End Property
Public Property Let CustomerCode(ByVal newval As String)
    If Trim(CustomerCode_) <> Trim(newval) Then
        CustomerCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerName() As String
    CustomerName = custName_
End Property
Public Property Let CustomerName(ByVal newval As String)
    If Trim(custName_) <> Trim(newval) Then
        custName_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CustomerNameCer() As String
    CustomerNameCer = custNameCer_
End Property
Public Property Let CustomerNameCer(ByVal newval As String)
    If Trim(custNameCer_) <> Trim(newval) Then
        custNameCer_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SubContact() As String
    SubContact = SubContact_
End Property
Public Property Let SubContact(ByVal newval As String)
    If Trim(SubContact_) <> Trim(newval) Then
        SubContact_ = Trim(newval)
    End If
End Property

Public Property Get CustomerShortName() As String
    CustomerShortName = custShortName_
End Property
Public Property Let CustomerShortName(ByVal newval As String)
    If Trim(custShortName_) <> Trim(newval) Then
        custShortName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SaleDistrictID() As String
    SaleDistrictID = saleDistrictID_
End Property
Public Property Let SaleDistrictID(ByVal newval As String)
    If Trim(saleDistrictID_) <> Trim(newval) Then
        saleDistrictID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SaleDistrictName() As String
    SaleDistrictName = saleDistrictName_
End Property
Public Property Let SaleDistrictName(ByVal newval As String)
    If Trim(saleDistrictName_) <> Trim(newval) Then
        saleDistrictName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SaleZoneID() As String
    SaleZoneID = saleZoneID_
End Property
Public Property Let SaleZoneID(ByVal newval As String)
    If Trim(saleZoneID_) <> Trim(newval) Then
        saleZoneID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SaleZoneName() As String
    SaleZoneName = saleZoneName_
End Property
Public Property Let SaleZoneName(ByVal newval As String)
    If Trim(saleZoneName_) <> Trim(newval) Then
        saleZoneName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustRemark() As String
    CustRemark = custRemark_
End Property
Public Property Let CustRemark(ByVal newval As String)
    If Trim(custRemark_) <> Trim(newval) Then
        custRemark_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BillingInfo() As BWGCustomerBillInfo
Dim tmpS As New BWGSearchManager
Dim tmpCol As Collection
    If billingInfo_ Is Nothing Then
        Set tmpCol = tmpS.CustomerBillingInfoSearch("CustomerID='" & id_ & "'")
        If tmpCol.Count > 0 Then
            Set billingInfo_ = tmpCol(1)
        Else
            Set billingInfo_ = New BWGCustomerBillInfo
            billingInfo_.ID = GetGUID
            billingInfo_.CustomerID = id_
            billingInfo_.CustomerName = custName_
        End If
    End If
    Set BillingInfo = billingInfo_
End Property

Public Property Set BillingInfo(newval As BWGCustomerBillInfo)
    Set billingInfo_ = newval
'    isChanged_ = True
End Property

Public Property Get CustomerType() As String
    CustomerType = custType_
End Property
Public Property Let CustomerType(ByVal newval As String)
    If Trim(custType_) <> Trim(newval) Then
        custType_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ParentCustID() As String
    ParentCustID = parentCustID_
End Property
Public Property Let ParentCustID(ByVal newval As String)
    If Trim(parentCustID_) <> Trim(newval) Then
        parentCustID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CreatedStaffID() As String
    CreatedStaffID = createdStaffID_
End Property
Public Property Let CreatedStaffID(ByVal newval As String)
    If Trim(createdStaffID_) <> Trim(newval) Then
        createdStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CreatedStaffName() As String
    CreatedStaffName = createdStaffName_
End Property
Public Property Let CreatedStaffName(ByVal newval As String)
    If Trim(createdStaffName_) <> Trim(newval) Then
        createdStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CreatedDate() As String
    CreatedDate = createdDate_
End Property
Public Property Let CreatedDate(ByVal newval As String)
    If Trim(createdDate_) <> Trim(newval) Then
        createdDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get UnderSaleStaffID() As String
    UnderSaleStaffID = underSaleStaffID_
End Property
Public Property Let UnderSaleStaffID(ByVal newval As String)
    If Trim(underSaleStaffID_) <> Trim(newval) Then
        underSaleStaffID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get UnderSaleName() As String
    UnderSaleName = underSaleName_
End Property
Public Property Let UnderSaleName(ByVal newval As String)
    If Trim(underSaleName_) <> Trim(newval) Then
        underSaleName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get UnderSaleTelNo() As String
    UnderSaleTelNo = underSaleTelNo_
End Property
Public Property Let UnderSaleTelNo(ByVal newval As String)
    If Trim(underSaleTelNo_) <> Trim(newval) Then
        underSaleTelNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustGroupID() As String
    CustGroupID = custGroupID_
End Property
Public Property Let CustGroupID(ByVal newval As String)
    If Trim(custGroupID_) <> Trim(newval) Then
        custGroupID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CustGroupName() As String
    CustGroupName = custGroupName_
End Property
Public Property Let CustGroupName(ByVal newval As String)
    If Trim(custGroupName_) <> Trim(newval) Then
        custGroupName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ParentCustCode() As String
    ParentCustCode = parentCustCode_
End Property
Public Property Let ParentCustCode(ByVal newval As String)
    If Trim(parentCustCode_) <> Trim(newval) Then
        parentCustCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ParentCustName() As String
    ParentCustName = parentCustName_
End Property
Public Property Let ParentCustName(ByVal newval As String)
    If Trim(parentCustName_) <> Trim(newval) Then
        parentCustName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Address() As String
    Address = CustAddr_
End Property
Public Property Let Address(ByVal newval As String)
    If Trim(CustAddr_) <> Trim(newval) Then
        CustAddr_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TumbolID() As String
    TumbolID = tumbolID_
End Property
Public Property Let TumbolID(ByVal newval As String)
    If Trim(tumbolID_) <> Trim(newval) Then
        tumbolID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TumbolName() As String
    TumbolName = tumbolName_
End Property
Public Property Let TumbolName(ByVal newval As String)
    If Trim(tumbolName_) <> Trim(newval) Then
        tumbolName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get AumphurID() As String
    AumphurID = aumphurID_
End Property
Public Property Let AumphurID(ByVal newval As String)
    If Trim(aumphurID_) <> Trim(newval) Then
        aumphurID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get AumphurName() As String
    AumphurName = aumphurName_
End Property
Public Property Let AumphurName(ByVal newval As String)
    If Trim(aumphurName_) <> Trim(newval) Then
        aumphurName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceID() As String
    ProvinceID = provinceID_
End Property
Public Property Let ProvinceID(ByVal newval As String)
    If Trim(provinceID_) <> Trim(newval) Then
        provinceID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceName() As String
    ProvinceName = provinceName_
End Property
Public Property Let ProvinceName(ByVal newval As String)
    If Trim(provinceName_) <> Trim(newval) Then
        provinceName_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CustdeptName() As String
    CustdeptName = CustdeptName_
End Property
Public Property Let CustdeptName(ByVal newval As String)
    If Trim(CustdeptName_) <> Trim(newval) Then
        CustdeptName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ZipCode() As String
    ZipCode = zipCode_
End Property
Public Property Let ZipCode(ByVal newval As String)
    If Trim(zipCode_) <> Trim(newval) Then
        zipCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Creditterm() As String
    Creditterm = Creditterm_
End Property
Public Property Let Creditterm(ByVal newval As String)
    If Trim(Creditterm_) <> Trim(newval) Then
        Creditterm_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Country() As String
    Country = Country_
End Property
Public Property Let Country(ByVal newval As String)
    If Trim(Country_) <> Trim(newval) Then
        Country_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TelNo() As String
    TelNo = TelNo_
End Property
Public Property Let TelNo(ByVal newval As String)
    If Trim(TelNo_) <> Trim(newval) Then
        TelNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FaxNo() As String
    FaxNo = faxNo_
End Property
Public Property Let FaxNo(ByVal newval As String)
    If Trim(faxNo_) <> Trim(newval) Then
        faxNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EmailAddr() As String
    EmailAddr = emailAddr_
End Property
Public Property Let EmailAddr(ByVal newval As String)
    If Trim(emailAddr_) <> Trim(newval) Then
        emailAddr_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Website() As String
    Website = website_
End Property
Public Property Let Website(ByVal newval As String)
    If Trim(website_) <> Trim(newval) Then
        website_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ContactPerson() As String
    ContactPerson = contactPerson_
End Property
Public Property Let ContactPerson(ByVal newval As String)
    If Trim(contactPerson_) <> Trim(newval) Then
        contactPerson_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get PositionName() As String
    PositionName = positionName_
End Property
Public Property Let PositionName(ByVal newval As String)
    If Trim(positionName_) <> Trim(newval) Then
        positionName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get DepartmentName() As String
    DepartmentName = deptName_
End Property
Public Property Let DepartmentName(ByVal newval As String)
    If Trim(deptName_) <> Trim(newval) Then
        deptName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get IndustryTypeID() As String
    IndustryTypeID = industryTypeID_
End Property
Public Property Let IndustryTypeID(ByVal newval As String)
    If Trim(industryTypeID_) <> Trim(newval) Then
        industryTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get IndustryTypeName() As String
    IndustryTypeName = industryTypeName_
End Property
Public Property Let IndustryTypeName(ByVal newval As String)
    If Trim(industryTypeName_) <> Trim(newval) Then
        industryTypeName_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get IndustryTypeIDNew() As String
    IndustryTypeIDNew = industryTypeIDNew_
End Property
Public Property Let IndustryTypeIDNew(ByVal newval As String)
    If Trim(industryTypeIDNew_) <> Trim(newval) Then
        industryTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TypeDesc() As String
    TypeDesc = TypeDesc_
End Property
Public Property Let TypeDesc(ByVal newval As String)
    If Trim(TypeDesc_) <> Trim(newval) Then
        TypeDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ZoneID() As String 'ปรับแก้จาก Zone เป็น Sale District
    ZoneID = zoneID_
End Property
Public Property Let ZoneID(ByVal newval As String)
    If Trim(zoneID_) <> Trim(newval) Then
        zoneID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ZoneName() As String
    ZoneName = zoneName_
End Property
Public Property Let ZoneName(ByVal newval As String)
    If Trim(zoneName_) <> Trim(newval) Then
        zoneName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get RegisterNo() As String
    RegisterNo = RegisterNo_
End Property
Public Property Let RegisterNo(ByVal newval As String)
    If Trim(RegisterNo_) <> Trim(newval) Then
        RegisterNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteManageID() As String
    WasteManageID = WasteManageID_
End Property
Public Property Let WasteManageID(ByVal newval As String)
    If Trim(WasteManageID_) <> Trim(newval) Then
        WasteManageID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TaxNo() As String
    TaxNo = taxNo_
End Property
Public Property Let TaxNo(ByVal newval As String)
    If Trim(taxNo_) <> Trim(newval) Then
        taxNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isAccountOk() As Boolean
    If Trim(isAccountOK_) = "Y" Then
        isAccountOk = True
    Else
        isAccountOk = False
    End If
End Property
Public Property Let isAccountOk(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isAccountOK_) <> Trim(tmpStr) Then
        isAccountOK_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get WasteItems() As Collection
Dim tmpS As New BWGSearchManager
    If colWasteItem_ Is Nothing Then Set colWasteItem_ = tmpS.WasteDataCRSearch("CustomerID='" & id_ & "'", "vw_WasteDataCRSearch2")
    Set WasteItems = colWasteItem_
    Set tmpS = Nothing
End Property
Public Property Set WasteItems(ByVal newval As Collection)
    Set colWasteItem_ = newval
    isChanged_ = True
End Property
Public Sub RequeryWasteItems()
    Set colWasteItem_ = Nothing
End Sub

Public Property Get SiteAddresses() As Collection
Dim tmpS As New BWGSearchManager
    If colSiteAddr_ Is Nothing Then Set colSiteAddr_ = tmpS.CustomerSiteSearch("CustomerID='" & id_ & "'")
    Set SiteAddresses = colSiteAddr_
    Set tmpS = Nothing
End Property
Public Property Set SiteAddresses(ByVal newval As Collection)
    Set colSiteAddr_ = newval
    isChanged_ = True
End Property
Public Sub RequerySiteAddress()
    Set colSiteAddr_ = Nothing
End Sub

Public Property Get BillAddresses() As Collection
Dim tmpS As New BWGSearchManager
    If colBillAddr_ Is Nothing Then Set colBillAddr_ = tmpS.CustomerBillingAddrSearch("CustomerID='" & id_ & "'")
    Set BillAddresses = colBillAddr_
    Set tmpS = Nothing
End Property

Public Property Set BillAddresses(ByVal newval As Collection)
    Set colBillAddr_ = newval
'    isChanged_ = True
End Property
Public Sub RequeryBillAddress()
    Set colBillAddr_ = Nothing
End Sub
Public Property Get Contracts() As Collection
Dim tmpS As New BWGSearchManager
    If colContract_ Is Nothing Then Set colContract_ = tmpS.ContractSearch("CustomerID='" & id_ & "'")
    Set Contracts = colContract_
    Set tmpS = Nothing
End Property
Public Property Set Contracts(ByVal newval As Collection)
    Set colContract_ = newval
    isChanged_ = True
End Property
Public Sub RequeryContracts()
    Set colContract_ = Nothing
End Sub

Public Property Get CmmQuo() As Collection
Dim tmpS As New BWGSearchManager
'    If CmmQuo_ Is Nothing Then
    Set CmmQuo_ = tmpS.CmmQuoSearch("CustomerID='" & id_ & "'")
    Set CmmQuo = CmmQuo_
    Set tmpS = Nothing
End Property
Public Property Set CmmQuo(ByVal newval As Collection)
    Set CmmQuo_ = newval
    isChanged_ = True
End Property

Public Sub RequeryCmmQuo()
    Set CmmQuo_ = Nothing
End Sub

Public Property Get CustStaffs() As Collection
Dim tmpS As New BWGSearchManager
'    If colCustStaff_ Is Nothing Then Set colCustStaff_ = tmpS.CustomerStaffSearch("CustomerID='" & id_ & "'")
    Set colCustStaff_ = tmpS.CustomerStaffSearch("CustomerID='" & id_ & "'")
    Set CustStaffs = colCustStaff_
    Set tmpS = Nothing
End Property
Public Property Set CustStaffs(ByVal newval As Collection)
    Set colCustStaff_ = newval
    isChanged_ = True
End Property
Public Sub RequeryCustStaff()
    Set colContract_ = Nothing
End Sub

Public Property Get OtherRemarks() As Collection
Dim tmpS As New BWGSearchManager
    If colOtherRemark_ Is Nothing Then Set colOtherRemark_ = tmpS.CustOrderRemarkSearch("CustomerID='" & id_ & "'")
    Set OtherRemarks = colOtherRemark_
    Set tmpS = Nothing
End Property
Public Property Set OtherRemarks(ByVal newval As Collection)
    Set colOtherRemark_ = newval
    isChanged_ = True
End Property
Public Sub RequeryOtherRemark()
    Set colOtherRemark_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Property Get Status() As String
    Status = Status_
End Property
Public Property Let Status(ByVal newval As String)
    If Trim(Status_) <> Trim(newval) Then
        Status_ = Trim(newval)
    End If
End Property

Public Property Get isNew() As Boolean
    If Trim(isNew_) = "Y" Then
        isNew = True
    Else
        isNew = False
    End If
End Property
Public Property Let isNew(ByVal newval As Boolean)
    If newval = True Then
        isNew_ = "Y"
    Else
        isNew_ = "N"
    End If
End Property

Public Property Get SaleAP() As String
    SaleAP = SaleAP_
End Property
Public Property Let SaleAP(ByVal newval As String)
    If Trim(SaleAP_) <> Trim(newval) Then
        SaleAP_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Sale2ID() As String
    Sale2ID = Sale2ID_
End Property
Public Property Let Sale2ID(ByVal newval As String)
    If Trim(Sale2ID_) <> Trim(newval) Then
        Sale2ID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Sale2Name() As String
    Sale2Name = Sale2Name_
End Property
Public Property Let Sale2Name(ByVal newval As String)
    If Trim(Sale2Name_) <> Trim(newval) Then
        Sale2Name_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Comment() As String
    Comment = Comment_
End Property
Public Property Let Comment(ByVal newval As String)
        Comment_ = Trim(newval)
        isChanged_ = True
End Property

Public Property Get CommentDate() As String
    CommentDate = CommentDate_
End Property
Public Property Let CommentDate(ByVal newval As String)
        CommentDate_ = Trim(newval)
        isChanged_ = True
End Property

Public Property Get CancelDate() As String
    CancelDate = cancelDate_
End Property
Public Property Let CancelDate(ByVal newval As String)
        cancelDate_ = Trim(newval)
End Property

Public Property Get CancelStaff() As String
    CancelStaff = CancelStaff_
End Property
Public Property Let CancelStaff(ByVal newval As String)
        CancelStaff_ = Trim(newval)
End Property

Public Property Get DTeam() As Boolean
    If Trim(DTeam_) = "Y" Then
        DTeam = True
    Else
        DTeam = False
    End If
End Property
Public Property Let DTeam(ByVal newval As Boolean)
    If newval = True Then
        DTeam_ = "Y"
    Else
        DTeam_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get LastUser() As String
    LastUser = LastUser_
End Property
Public Property Let LastUser(ByVal newval As String)
    If Trim(LastUser_) <> Trim(newval) Then
        LastUser_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LastUpdate() As String
    LastUpdate = LastUpdate_
End Property
Public Property Let LastUpdate(ByVal newval As String)
    If Trim(LastUpdate_) <> Trim(newval) Then
        LastUpdate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ischeck() As String
      ischeck = ischeck_
End Property
Public Property Let ischeck(ByVal newval As String)
    If Trim(ischeck_) <> Trim(newval) Then
        ischeck_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isVIPCR() As String
      isVIPCR = isVIPCR_
End Property
Public Property Let isVIPCR(ByVal newval As String)
    If Trim(isVIPCR_) <> Trim(newval) Then
        isVIPCR_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LastUserTR() As String
    LastUserTR = LastUserTR_
End Property
Public Property Let LastUserTR(ByVal newval As String)
    If Trim(LastUserTR_) <> Trim(newval) Then
        LastUserTR_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LastUpdateTR() As String
    LastUpdateTR = LastUpdateTR_
End Property
Public Property Let LastUpdateTR(ByVal newval As String)
    If Trim(LastUpdateTR_) <> Trim(newval) Then
        LastUpdateTR_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CommentTimeNo() As String
    CommentTimeNo = CommentTimeNo_
End Property
Public Property Let CommentTimeNo(ByVal newval As String)
    If Trim(CommentTimeNo_) <> Trim(newval) Then
        CommentTimeNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isWeigth() As String
    If Trim(isWeigth_) = "Y" Then
        isWeigth = "Y"
    Else
        isWeigth = "N"
    End If
End Property
Public Property Let isWeigth(ByVal newval As String)
    If newval = True Then
        isWeigth_ = "Y"
    Else
        isWeigth_ = "N"
    End If
End Property


' ข้อมูล Box

Public Property Get isBox() As String
      isBox = isBox_
End Property
Public Property Let isBox(ByVal newval As String)
    If Trim(isBox_) <> Trim(newval) Then
        isBox_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get isLUGGERBOX() As String
      isLUGGERBOX = isLUGGERBOX_
End Property
Public Property Let isLUGGERBOX(ByVal newval As String)
    If Trim(isLUGGERBOX_) <> Trim(newval) Then
        isLUGGERBOX_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get LUGBOX() As String
      LUGBOX = LUGBOX_
End Property
Public Property Let LUGBOX(ByVal newval As String)
    If Trim(LUGBOX_) <> Trim(newval) Then
        LUGBOX_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get LUGBOXFrom() As String
      LUGBOXFrom = LUGBOXFrom_
End Property
Public Property Let LUGBOXFrom(ByVal newval As String)
    If Trim(LUGBOXFrom_) <> Trim(newval) Then
        LUGBOXFrom_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get LUGBOXBack() As String
      LUGBOXBack = LUGBOXBack_
End Property
Public Property Let LUGBOXBack(ByVal newval As String)
    If Trim(LUGBOXBack_) <> Trim(newval) Then
        LUGBOXBack_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get LUGBOXBackV() As String
      LUGBOXBackV = LUGBOXBackV_
End Property
Public Property Let LUGBOXBackV(ByVal newval As String)
    If Trim(LUGBOXBackV_) <> Trim(newval) Then
        LUGBOXBackV_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get LUGBoxW() As String
      LUGBoxW = LUGBoxW_
End Property
Public Property Let LUGBoxW(ByVal newval As String)
    If Trim(LUGBoxW_) <> Trim(newval) Then
        LUGBoxW_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get isROLGERBOX() As String
      isROLGERBOX = isROLGERBOX_
End Property
Public Property Let isROLGERBOX(ByVal newval As String)
    If Trim(isROLGERBOX_) <> Trim(newval) Then
        isROLGERBOX_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get ROLBOX() As String
      ROLBOX = ROLBOX_
End Property
Public Property Let ROLBOX(ByVal newval As String)
    If Trim(ROLBOX_) <> Trim(newval) Then
        ROLBOX_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get ROLBOXFrom() As String
      ROLBOXFrom = ROLBOXFrom_
End Property
Public Property Let ROLBOXFrom(ByVal newval As String)
    If Trim(ROLBOXFrom_) <> Trim(newval) Then
        ROLBOXFrom_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get ROLBOXBack() As String
      ROLBOXBack = ROLBOXBack_
End Property
Public Property Let ROLBOXBack(ByVal newval As String)
    If Trim(ROLBOXBack_) <> Trim(newval) Then
        ROLBOXBack_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get ROLBOXBackv() As String
      ROLBOXBackv = ROLBOXBackv_
End Property
Public Property Let ROLBOXBackv(ByVal newval As String)
    If Trim(ROLBOXBackv_) <> Trim(newval) Then
        ROLBOXBackv_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get ROLBoxW() As String
      ROLBoxW = ROLBoxW_
End Property
Public Property Let ROLBoxW(ByVal newval As String)
    If Trim(ROLBoxW_) <> Trim(newval) Then
        ROLBoxW_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get isOneTBOX() As String
      isOneTBOX = isOneTBOX_
End Property
Public Property Let isOneTBOX(ByVal newval As String)
    If Trim(isOneTBOX_) <> Trim(newval) Then
        isOneTBOX_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get OneT() As String
      OneT = OneT_
End Property
Public Property Let OneT(ByVal newval As String)
    If Trim(OneT_) <> Trim(newval) Then
        OneT_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get OneTFrom() As String
      OneTFrom = OneTFrom_
End Property
Public Property Let OneTFrom(ByVal newval As String)
    If Trim(OneTFrom_) <> Trim(newval) Then
        OneTFrom_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get OneTBack() As String
      OneTBack = OneTBack_
End Property
Public Property Let OneTBack(ByVal newval As String)
    If Trim(OneTBack_) <> Trim(newval) Then
        OneTBack_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get OneTBackv() As String
      OneTBackv = OneTBackv_
End Property
Public Property Let OneTBackv(ByVal newval As String)
    If Trim(OneTBackv_) <> Trim(newval) Then
        OneTBackv_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get OneTBoxW() As String
      OneTBoxW = OneTBoxW_
End Property
Public Property Let OneTBoxW(ByVal newval As String)
    If Trim(OneTBoxW_) <> Trim(newval) Then
        OneTBoxW_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get isTwoTBOX() As String
      isTwoTBOX = isTwoTBOX_
End Property
Public Property Let isTwoTBOX(ByVal newval As String)
    If Trim(isTwoTBOX_) <> Trim(newval) Then
        isTwoTBOX_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TwoT() As String
      TwoT = TwoT_
End Property
Public Property Let TwoT(ByVal newval As String)
    If Trim(TwoT_) <> Trim(newval) Then
        TwoT_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TwoTFrom() As String
      TwoTFrom = TwoTFrom_
End Property
Public Property Let TwoTFrom(ByVal newval As String)
    If Trim(TwoTFrom_) <> Trim(newval) Then
        TwoTFrom_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TwoTBack() As String
      TwoTBack = TwoTBack_
End Property
Public Property Let TwoTBack(ByVal newval As String)
    If Trim(TwoTBack_) <> Trim(newval) Then
        TwoTBack_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TwoTBackv() As String
      TwoTBackv = TwoTBackv_
End Property
Public Property Let TwoTBackv(ByVal newval As String)
    If Trim(TwoTBackv_) <> Trim(newval) Then
        TwoTBackv_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TwoTBoxW() As String
      TwoTBoxW = TwoTBoxW_
End Property
Public Property Let TwoTBoxW(ByVal newval As String)
    If Trim(TwoTBoxW_) <> Trim(newval) Then
        TwoTBoxW_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isTruckBOX() As String
      isTruckBOX = isTruckBOX_
End Property
Public Property Let isTruckBOX(ByVal newval As String)
    If Trim(isTruckBOX_) <> Trim(newval) Then
        isTruckBOX_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TruckBOX() As String
      TruckBOX = TruckBOX_
End Property
Public Property Let TruckBOX(ByVal newval As String)
    If Trim(TruckBOX_) <> Trim(newval) Then
        TruckBOX_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TruckBOXFrom() As String
      TruckBOXFrom = TruckBOXFrom_
End Property
Public Property Let TruckBOXFrom(ByVal newval As String)
    If Trim(TruckBOXFrom_) <> Trim(newval) Then
        TruckBOXFrom_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TruckBOXBack() As String
      TruckBOXBack = TruckBOXBack_
End Property
Public Property Let TruckBOXBack(ByVal newval As String)
    If Trim(TruckBOXBack_) <> Trim(newval) Then
        TruckBOXBack_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TruckBOXBackv() As String
      TruckBOXBackv = TruckBOXBackv_
End Property
Public Property Let TruckBOXBackv(ByVal newval As String)
    If Trim(TruckBOXBackv_) <> Trim(newval) Then
        TruckBOXBackv_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TruckBOXW() As String
      TruckBOXW = TruckBOXW_
End Property
Public Property Let TruckBOXW(ByVal newval As String)
    If Trim(TruckBOXW_) <> Trim(newval) Then
        TruckBOXW_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BoxDesc() As String
      BoxDesc = BoxDesc_
End Property
Public Property Let BoxDesc(ByVal newval As String)
    If Trim(BoxDesc_) <> Trim(newval) Then
        BoxDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isHaz() As String
    isHaz = isHaz_
End Property
Public Property Let isHaz(ByVal newval As String)
    If Trim(isHaz_) <> Trim(newval) Then
        isHaz_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get isNon() As String
    isNon = isNon_
End Property
Public Property Let isNon(ByVal newval As String)
    If Trim(isNon_) <> Trim(newval) Then
        isNon_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get PriceBox() As String
    PriceBox = PriceBox_
End Property
Public Property Let PriceBox(ByVal newval As String)
    If Trim(PriceBox_) <> Trim(newval) Then
        PriceBox_ = Trim(newval)
        isChanged_ = True
    End If
    
End Property
Public Property Get StpTrn() As String
    StpTrn = StpTrn_
End Property
Public Property Let StpTrn(ByVal newval As String)
    If Trim(StpTrn_) <> Trim(newval) Then
        StpTrn = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Lack() As String
      Lack = Lack_
End Property
Public Property Let Lack(ByVal newval As String)
    If Trim(Lack_) <> Trim(newval) Then
        Lack_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get LackBack() As String
      LackBack = LackBack_
End Property
Public Property Let LackBack(ByVal newval As String)
    If Trim(LackBack_) <> Trim(newval) Then
        LackBack_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get LackFrom() As String
      LackFrom = LackFrom_
End Property
Public Property Let LackFrom(ByVal newval As String)
    If Trim(LackFrom_) <> Trim(newval) Then
        LackFrom_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get LackTo() As String
      LackTo = LackTo_
End Property
Public Property Let LackTo(ByVal newval As String)
    If Trim(LackTo_) <> Trim(newval) Then
        LackTo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get LackW() As String
      LackW = LackW_
End Property
Public Property Let LackW(ByVal newval As String)
    If Trim(LackW_) <> Trim(newval) Then
        LackW_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BoutS() As String
      BoutS = BoutS_
End Property
Public Property Let BoutS(ByVal newval As String)
    If Trim(BoutS_) <> Trim(newval) Then
        BoutS_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BoutSBack() As String
      BoutSBack = BoutSBack_
End Property
Public Property Let BoutSBack(ByVal newval As String)
    If Trim(BoutSBack_) <> Trim(newval) Then
        BoutSBack_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BoutSFrom() As String
      BoutSFrom = BoutSFrom_
End Property
Public Property Let BoutSFrom(ByVal newval As String)
    If Trim(BoutSFrom_) <> Trim(newval) Then
        BoutSFrom_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BoutSTo() As String
      BoutSTo = BoutSTo_
End Property
Public Property Let BoutSTo(ByVal newval As String)
    If Trim(BoutSTo_) <> Trim(newval) Then
        BoutSTo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BoutSW() As String
      BoutSW = BoutSW_
End Property
Public Property Let BoutSW(ByVal newval As String)
    If Trim(BoutSW_) <> Trim(newval) Then
        BoutSW_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WaterBout() As String
      WaterBout = WaterBout_
End Property
Public Property Let WaterBout(ByVal newval As String)
    If Trim(WaterBout_) <> Trim(newval) Then
        WaterBout_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get WaterBoutBack() As String
      WaterBoutBack = WaterBoutBack_
End Property
Public Property Let WaterBoutBack(ByVal newval As String)
    If Trim(WaterBoutBack_) <> Trim(newval) Then
        WaterBoutBack_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get WaterBoutFrom() As String
      WaterBoutFrom = WaterBoutFrom_
End Property
Public Property Let WaterBoutFrom(ByVal newval As String)
    If Trim(WaterBoutFrom_) <> Trim(newval) Then
        WaterBoutFrom_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get WaterBoutTo() As String
      WaterBoutTo = WaterBoutTo_
End Property
Public Property Let WaterBoutTo(ByVal newval As String)
    If Trim(WaterBoutTo_) <> Trim(newval) Then
        WaterBoutTo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get WaterBoutW() As String
      WaterBoutW = WaterBoutW_
End Property
Public Property Let WaterBoutW(ByVal newval As String)
    If Trim(WaterBoutW_) <> Trim(newval) Then
        WaterBoutW_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BoutCage() As String
      BoutCage = BoutCage_
End Property
Public Property Let BoutCage(ByVal newval As String)
    If Trim(BoutCage_) <> Trim(newval) Then
        BoutCage_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BoutCageBack() As String
      BoutCageBack = BoutCageBack_
End Property
Public Property Let BoutCageBack(ByVal newval As String)
    If Trim(BoutCageBack_) <> Trim(newval) Then
        BoutCageBack_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BoutCageFrom() As String
      BoutCageFrom = BoutCageFrom_
End Property
Public Property Let BoutCageFrom(ByVal newval As String)
    If Trim(BoutCageFrom_) <> Trim(newval) Then
        BoutCageFrom_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BoutCageTo() As String
      BoutCageTo = BoutCageTo_
End Property
Public Property Let BoutCageTo(ByVal newval As String)
    If Trim(BoutCageTo_) <> Trim(newval) Then
        BoutCageTo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get BoutCageW() As String
      BoutCageW = BoutCageW_
End Property
Public Property Let BoutCageW(ByVal newval As String)
    If Trim(BoutCageW_) <> Trim(newval) Then
        BoutCageW_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Sub Save()
'    On Error GoTo ErrD
    Dim tmpRec As ADODB.Recordset
    Dim tmpSite As BWGCustomerSite
    Dim tmpBillAddr As BWGCustomerBillAddr
    Dim sqlStr As String
    Dim UserGrp As String
    Dim i As Integer
    Dim CanUpdate As Boolean
    Dim tmpStr As String
    
    Set tmpRec = New ADODB.Recordset
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbCustomer WHERE CustomerID='" & id_ & "'"
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("CustomerID") = id_
        tmpRec("CreatedByStaffID") = Trim(createdStaffID_)
        tmpRec("CreatedDate") = Trim(createdDate_)
        Set tmpSite = New BWGCustomerSite
        With tmpSite
            .ID = GetGUID
            .Address = Trim(CustAddr_)
            .SiteName = "ที่ตั้งโรงงาน"
            .TumbolID = Trim(tumbolID_)
            .ZipCode = Trim(zipCode_)
            .Country = Trim(Country_)
            .CustomerID = id_
            .TelNo = Trim(TelNo_)
            .FaxNo = Trim(faxNo_)
        End With
        Set colSiteAddr_ = Nothing
        Set tmpBillAddr = New BWGCustomerBillAddr
        With tmpBillAddr
            .ID = GetGUID
            .BillName = "ที่ตั้งโรงงาน"
            .CustomerID = id_
            .FaxNo = Trim(faxNo_)
            .TelNo = Trim(TelNo_)
            .TumbolID = Trim(tumbolID_)
            .ZipCode = Trim(zipCode_)
            .Country = Trim(Country_)
        End With
        Set colBillAddr_ = Nothing
    End If
    tmpRec("CustRemark") = Trim(custRemark_)
    tmpRec("CompanyName") = Trim(custName_)
    tmpRec("CustShortName") = Trim(custShortName_)
    tmpRec("Addr1") = Trim(CustAddr_)
    tmpRec("TumbolID") = Trim(tumbolID_)
    tmpRec("ZipCode") = Trim(zipCode_)
    tmpRec("Country") = Trim(Country_)
    tmpRec("TelNo") = Trim(TelNo_)
    tmpRec("FaxNo") = Trim(faxNo_)
    tmpRec("WebSite") = Trim(website_)
    tmpRec("EmailAddr") = Trim(emailAddr_)
    tmpRec("ContactPerson") = Trim(contactPerson_)
    tmpRec("PositionName") = Trim(positionName_)
    tmpRec("ContactDeptNo") = Trim(deptName_)
    tmpRec("IndustryTypeID") = Trim(industryTypeID_)
    tmpRec("ZoneID") = Trim(zoneID_)
    tmpRec("RegisterNo") = Trim(RegisterNo_)
    tmpRec("14Digit") = Trim(Digit14_)
    tmpRec("WasteManageID") = Trim(WasteManageID_)
    tmpRec("TaxNo") = Trim(taxNo_)
    tmpRec("isAccountantOK") = Trim(isAccountOK_)
    tmpRec("CustTypeFlag") = Trim(custType_)
    tmpRec("ParentCustID") = Trim(parentCustID_)
    tmpRec("UnderSaleStaffID") = Trim(underSaleStaffID_)
    tmpRec("CustGroupID") = Trim(custGroupID_)
    tmpRec("SaleDistrictID") = Trim(saleDistrictID_)
    tmpRec("APSale") = Trim(SaleAP_)
    tmpRec("Sale2ID") = Trim(Sale2ID_)
    tmpRec("Sale2Name") = Trim(Sale2Name_)
    tmpRec("Comment") = Trim(Comment_)
    tmpRec("CommentDate") = Trim(CommentDate_)
    tmpRec!DTeam = Trim("" & DTeam_)
    tmpRec!ischeck = Trim("" & ischeck_)
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    tmpRec("CommentTimeNo") = Trim(CommentTimeNo_)
    tmpRec("isWeigth") = Trim(isWeigth_)
    tmpRec("RefCode") = Trim(Refcode_)
    tmpRec("isBox") = Trim(isBox_)
    tmpRec("isLUGGERBOX") = Trim(isLUGGERBOX_)
    tmpRec("LUGBOX") = Trim(LUGBOX_)
    tmpRec("LUGBOXFrom") = Trim(LUGBOXFrom_)
    tmpRec("LUGBOXBack") = Trim(LUGBOXBack_)
    tmpRec("LUGBOXBackV") = Trim(LUGBOXBackV_)
    tmpRec("LUGBoxW") = Trim(LUGBoxW_)
    tmpRec("isROLGERBOX") = Trim(isROLGERBOX_)
    tmpRec("ROLBOX") = Trim(ROLBOX_)
    tmpRec("ROLBOXFrom") = Trim(ROLBOXFrom_)
    tmpRec("ROLBOXBack") = Trim(ROLBOXBack_)
    tmpRec("ROLBOXBackv") = Trim(ROLBOXBackv_)
    tmpRec("ROLBoxW") = Trim(ROLBoxW_)
    tmpRec("isOneTBOX") = Trim(isOneTBOX_)
    tmpRec("OneT") = Trim(OneT_)
    tmpRec("OneTFrom") = Trim(OneTFrom_)
    tmpRec("OneTBack") = Trim(OneTBack_)
    tmpRec("OneTBackv") = Trim(OneTBackv_)
    tmpRec("OneTBoxW") = Trim(OneTBoxW_)
    tmpRec("isTwoTBOX") = Trim(isTwoTBOX_)
    tmpRec("TwoT") = Trim(TwoT_)
    tmpRec("TwoTFrom") = Trim(TwoTFrom_)
    tmpRec("TwoTBack") = Trim(TwoTBack_)
    tmpRec("TwoTBackv") = Trim(TwoTBackv_)
    tmpRec("TwoTBoxW") = Trim(TwoTBoxW_)
    tmpRec("isTruckBOX") = Trim(isTruckBOX_)
    tmpRec("TruckBOX") = Trim(TruckBOX_)
    tmpRec("TruckBOXFrom") = Trim(TruckBOXFrom_)
    tmpRec("TruckBOXBack") = Trim(TruckBOXBack_)
    tmpRec("TruckBOXBackv") = Trim(TruckBOXBackv_)
    tmpRec("TruckBOXW") = Trim(TruckBOXW_)
    tmpRec("BoxDesc") = Trim(BoxDesc_)
    tmpRec("isHaz") = Trim(isHaz_)
    tmpRec("isNon") = Trim(isNon_)
    tmpRec("creditterm") = Trim(Creditterm_)
    tmpRec("PriceBox") = Trim(PriceBox_)
    tmpRec("begindate") = Trim(BeginDate_)
    tmpRec("GPSRequest") = Trim(GPSRequest_)
    tmpRec("OverWeight") = Trim(OverWeight_)
    tmpRec("WasteSK2Exp") = Trim(WasteSK2Exp_)
    tmpRec("isVIPCR") = Trim(isVIPCR_)
    tmpRec("CustDept") = CustdeptName_
    tmpRec("CRStaffID") = CRStaffID_
    tmpRec("Lack") = Trim(Lack_)
    tmpRec("LackFrom") = Trim(LackFrom_)
    tmpRec("LackW") = Trim(LackW_)
    tmpRec("LackBack") = Trim(LackBack_)
    tmpRec("LackTo") = Trim(LackTo_)
    tmpRec("BoutS") = Trim(BoutS_)
    tmpRec("BoutsFrom") = Trim(BoutSFrom_)
    tmpRec("BoutsW") = Trim(BoutSW_)
    tmpRec("BoutsBack") = Trim(BoutSBack_)
    tmpRec("BoutsTo") = Trim(BoutSTo_)
    tmpRec("WaterBout") = Trim(WaterBout_)
    tmpRec("WaterBoutFrom") = Trim(WaterBoutFrom_)
    tmpRec("WaterBoutW") = Trim(WaterBoutW_)
    tmpRec("WaterBoutBack") = Trim(WaterBoutBack_)
    tmpRec("WaterBoutTo") = Trim(WaterBoutTo_)
    tmpRec("BoutCage") = Trim(BoutCage_)
    tmpRec("BoutCageFrom") = Trim(BoutCageFrom_)
    tmpRec("BoutCageW") = Trim(BoutCageW_)
    tmpRec("BoutCageBack") = Trim(BoutCageBack_)
    tmpRec("BoutCageTo") = Trim(BoutCageTo_)
    tmpRec("PicOfDay") = Trim(PicOfDay_)

    tmpRec.Update
    tmpRec.Requery
    CustomerCode_ = Trim("" & tmpRec("CustomerCode"))
    
    If Not tmpSite Is Nothing Then tmpSite.Save
    If Not tmpBillAddr Is Nothing Then tmpBillAddr.Save
    
    Set tmpSite = Nothing
    Set tmpBillAddr = Nothing
    Set tmpRec = Nothing
    
    isChanged_ = False
    CloseDBConnection
    Exit Sub
'ErrD:
'    GetMsg "เกิดข้อผิดพลาดขณะบันทึกข้อมูลลูกค้า", vbCritical
'    Exit Sub
End Sub

Public Sub UpdateTR()
    On Error GoTo ErrD
    Dim tmpRec As ADODB.Recordset
    Dim sqlStr As String
        
    Set tmpRec = New ADODB.Recordset
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbCustomer WHERE CustomerID='" & id_ & "'"
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If Not tmpRec.EOF Then
        tmpRec!LastUserTR = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
        tmpRec!LastUpdateTR = tmpRec!GetDate
        tmpRec.Update
    End If
    
    Set tmpRec = Nothing
    isChanged_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "เกิดข้อผิดพลาดขณะบันทึกข้อมูลลูกค้า", vbCritical
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbCustomer WHERE CustomerID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Function Cancel() As String
    On Error GoTo ErrD:
    Dim rs As New ADODB.Recordset
    Dim tSql As String
    Dim ToDate As Date
    tSql = "SELECT TOP(10) CustomerID FROM tbMenifestHeader WHERE CustomerID IN ('" & id_ & "')"
    Set rs = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If rs.RecordCount > 0 Then
        Cancel = "ลูกค้ารายนี้ (" & CustomerCode_ & ") มีการขนแล้ว ไม่สามารถลบได้ กรุณาตรวจสอบข้อมูลอีกครั้ง"
    Else
        DBConnExc "UPDATE tbCustomer  SET Status = 'C', CancelDate = getdate(), CancelStaff = '" & Trim(CurrentUser.StaffFName & " " & CurrentUser.StaffLName) & "' WHERE CustomerID='" & id_ & "'"
        Cancel = ""
    End If
    Exit Function
ErrD:
    Cancel = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("CustomerID"))
    CustomerCode_ = Trim("" & selRec("CustomerCode"))
    custName_ = Trim("" & selRec("CompanyName"))
    custNameCer_ = Trim("" & selRec("CompanyNameCer"))
    custShortName_ = Trim("" & selRec("CustShortName"))
    CustAddr_ = Trim("" & selRec("Addr1"))
    tumbolID_ = Trim("" & selRec("TumbolID"))
    tumbolName_ = Trim("" & selRec("TumbolName"))
    aumphurID_ = Trim("" & selRec("AumphurID"))
    aumphurName_ = Trim("" & selRec("AumphurName"))
    provinceID_ = Trim("" & selRec("ProvinceID"))
    provinceName_ = Trim("" & selRec("ProvinceName"))
    zipCode_ = Trim("" & selRec("ZipCode"))
    Country_ = Trim("" & selRec("Country"))
    TelNo_ = Trim("" & selRec("TelNo"))
    WasteSK2Exp_ = IIf(IsNull(selRec!WasteSK2Exp), "", selRec!WasteSK2Exp)
    faxNo_ = Trim("" & selRec("FaxNo"))
    website_ = Trim("" & selRec("WebSite"))
    emailAddr_ = Trim("" & selRec("EmailAddr"))
    contactPerson_ = Trim("" & selRec("ContactPerson"))
    positionName_ = Trim("" & selRec("PositionName"))
    deptName_ = Trim("" & selRec("ContactDeptNo"))
    industryTypeID_ = Trim("" & selRec("IndustryTypeID"))
    industryTypeName_ = Trim("" & selRec("TypeName"))
    industryTypeIDNew_ = Trim("" & selRec("IndustryTypeIDnew"))
    TypeDesc_ = Trim("" & selRec("TypeDesc"))
    zoneID_ = Trim("" & selRec("ZoneID"))
    zoneName_ = Trim("" & selRec("ZoneName"))
    RegisterNo_ = Trim("" & selRec("RegisterNo"))
    Digit14_ = Trim("" & selRec("14Digit"))
    WasteManageID_ = Trim("" & selRec("WasteManageID"))
    taxNo_ = Trim("" & selRec("TaxNo"))
    isAccountOK_ = Trim("" & selRec("isAccountantOK"))
    custType_ = Trim("" & selRec("CustTypeFlag"))
    parentCustID_ = Trim("" & selRec("ParentCustID"))
    parentCustCode_ = Trim("" & selRec("ParentCustCode"))
    parentCustName_ = Trim("" & selRec("ParentCustName"))
    createdStaffID_ = Trim("" & selRec("CreatedByStaffID"))
    createdDate_ = Trim("" & selRec("CreatedDate"))
    underSaleStaffID_ = Trim("" & selRec("UnderSaleStaffID"))
    underSaleName_ = Trim("" & selRec("SaleStaffFName")) & " " & Trim("" & selRec("SaleStaffLName")) & " (" & Trim("" & selRec("StaffCode")) & ")"
    custGroupID_ = Trim("" & selRec("CustGroupID"))
    custGroupName_ = Trim("" & selRec("CustGroupName"))
    createdStaffName_ = Trim("" & selRec("CreatedStaffFName")) & " " & Trim("" & selRec("CreatedStaffLName"))
    saleDistrictID_ = Trim("" & selRec("SaleDistrictID"))
    saleDistrictName_ = Trim("" & selRec("SDName"))
    saleZoneID_ = Trim("" & selRec("SaleZoneID"))
    saleZoneName_ = Trim("" & selRec("SaleZoneName"))
    custRemark_ = Trim("" & selRec("CustRemark"))
    underSaleTelNo_ = Trim("" & selRec("SaleTelNo"))
    Status_ = Trim("" & selRec("Status"))
    ischeck_ = Trim("" & selRec("ischeck"))
    isNew_ = IIf(IsNull(selRec!isNew), "N", selRec!isNew)
    Comment_ = Trim("" & selRec("Comment"))
    CommentDate_ = Trim("" & selRec("CommentDate"))
    CommentTimeNo_ = Trim("" & selRec("CommentTimeNo"))
    If isExist(selRec, "APSale") Then SaleAP_ = IIf(IsNull(selRec!APSale), "", selRec!APSale)
    If isExist(selRec, "Sale2ID") Then Sale2ID_ = IIf(IsNull(selRec!Sale2ID), "", selRec!Sale2ID)
    If isExist(selRec, "Sale2Name") Then Sale2Name_ = IIf(IsNull(selRec!Sale2Name), "", selRec!Sale2Name)
    If isExist(selRec, "CancelDate") Then cancelDate_ = IIf(IsNull(selRec!CancelDate), "", selRec!CancelDate)
    If isExist(selRec, "CancelStaff") Then CancelStaff_ = IIf(IsNull(selRec!CancelStaff), "", selRec!CancelStaff)
    If isExist(selRec, "SubContact") Then SubContact_ = IIf(IsNull(selRec!SubContact), "", selRec!SubContact)
    If isExist(selRec, "DTeam") Then DTeam_ = IIf(IsNull(selRec!DTeam), "N", selRec!DTeam)
    If isExist(selRec, "LastUser") Then LastUser_ = Trim("" & selRec!LastUser)
    If isExist(selRec, "LastUpdate") Then LastUpdate_ = Trim("" & selRec!LastUpdate)
    If isExist(selRec, "LastUserTR") Then LastUserTR_ = Trim("" & selRec!LastUserTR)
    If isExist(selRec, "LastUpdateTR") Then LastUpdateTR_ = Trim("" & selRec!LastUpdateTR)
    isWeigth_ = Trim("" & selRec("isWeigth"))
    
    isBox_ = IIf(IsNull(selRec!isBox), "N", selRec!isBox)
    isLUGGERBOX_ = IIf(IsNull(selRec!isLUGGERBOX), "N", selRec!isLUGGERBOX)
    LUGBOX_ = Trim("" & selRec("LUGBOX"))
    LUGBOXFrom_ = Trim("" & selRec("LUGBOXFrom"))
    LUGBOXBack_ = Trim("" & selRec("LUGBOXBack"))
    LUGBOXBackV_ = Trim("" & selRec("LUGBOXBackV"))
    LUGBoxW_ = Trim("" & selRec("LUGBoxW"))
    isROLGERBOX_ = IIf(IsNull(selRec!isROLGERBOX), "N", selRec!isROLGERBOX)
    ROLBOX_ = Trim("" & selRec("ROLBOX"))
    ROLBOXFrom_ = Trim("" & selRec("ROLBOXFrom"))
    ROLBOXBack_ = Trim("" & selRec("ROLBOXBack"))
    ROLBOXBackv_ = Trim("" & selRec("ROLBOXBackv"))
    ROLBoxW_ = Trim("" & selRec("ROLBoxW"))
    isOneTBOX_ = IIf(IsNull(selRec!isOneTBOX), "N", selRec!isOneTBOX)
    OneT_ = Trim("" & selRec("OneT"))
    OneTFrom_ = Trim("" & selRec("OneTFrom"))
    OneTBack_ = Trim("" & selRec("OneTBack"))
    OneTBackv_ = Trim("" & selRec("OneTBackv"))
    OneTBoxW_ = Trim("" & selRec("OneTBoxW"))
    isTwoTBOX_ = IIf(IsNull(selRec!isTwoTBOX), "N", selRec!isTwoTBOX)
    TwoT_ = Trim("" & selRec("TwoT"))
    TwoTFrom_ = Trim("" & selRec("TwoTFrom"))
    TwoTBack_ = Trim("" & selRec("TwoTBack"))
    TwoTBackv_ = Trim("" & selRec("TwoTBackv"))
    TwoTBoxW_ = Trim("" & selRec("TwoTBoxW"))
    isTruckBOX_ = IIf(IsNull(selRec!isTruckBOX), "N", selRec!isTruckBOX)
    TruckBOX_ = Trim("" & selRec("TruckBOX"))
    TruckBOXFrom_ = Trim("" & selRec("TruckBOXFrom"))
    TruckBOXBack_ = Trim("" & selRec("TruckBOXBack"))
    TruckBOXBackv_ = Trim("" & selRec("TruckBOXBackv"))
    TruckBOXW_ = Trim("" & selRec("TruckBOXW"))
    BoxDesc_ = Trim("" & selRec("BoxDesc"))
    Refcode_ = Trim("" & selRec("RefCode"))
    isHaz_ = IIf(IsNull(selRec!isHaz), "N", selRec!isHaz)
    isNon_ = IIf(IsNull(selRec!isNon), "N", selRec!isNon)
'    BeginDate_ = Trim("" & selRec("begindate"))
    PriceBox_ = Trim("" & selRec("PriceBox"))
    GPSRequest_ = Trim("" & selRec("GPSRequest"))
    OverWeight_ = Trim("" & selRec("OverWeight"))
    StpTrn_ = Trim("" & selRec("StopTrn"))
    isVIPCR_ = Trim("" & selRec("isVIPCR"))
    If IsNumeric(selRec("PicOfDay")) Then
     PicOfDay_ = CInt(selRec("PicOfDay"))
    Else
     PicOfDay_ = 0
    End If
   
    If isExist(selRec, "CustDept") Then
        CustdeptName_ = IIf(IsNull(selRec!CustDept), "", selRec!CustDept)
    End If
    If isExist(selRec, "creditterm") Then
        Creditterm_ = IIf(IsNull(Trim(selRec!Creditterm)), "0", selRec!Creditterm)
    End If
    If isExist(selRec, "CRStaffID") Then CRStaffID_ = Trim("" & selRec!CRStaffID)
    If isExist(selRec, "CRStaffName") Then CRStaffName_ = Trim("" & selRec!CRStaffName)
    If isExist(selRec, "Lack") Then Lack_ = Trim("" & selRec!Lack)
    If isExist(selRec, "LackFrom") Then LackFrom_ = Trim("" & selRec!LackFrom)
    If isExist(selRec, "LackW") Then LackW_ = Trim("" & selRec!LackW)
    If isExist(selRec, "LackBack") Then LackBack_ = Trim("" & selRec!LackBack)
    If isExist(selRec, "LackTo") Then LackTo_ = Trim("" & selRec!LackTo)
    If isExist(selRec, "BoutS") Then BoutS_ = Trim("" & selRec!BoutS)
    If isExist(selRec, "BoutsFrom") Then BoutSFrom_ = Trim("" & selRec!BoutSFrom)
    If isExist(selRec, "BoutsW") Then BoutSW_ = Trim("" & selRec!BoutSW)
    If isExist(selRec, "BoutsBack") Then BoutSBack_ = Trim("" & selRec!BoutSBack)
    If isExist(selRec, "BoutsTo") Then BoutSTo_ = Trim("" & selRec!BoutSTo)
    If isExist(selRec, "WaterBout") Then WaterBout_ = Trim("" & selRec!WaterBout)
    If isExist(selRec, "WaterBoutFrom") Then WaterBoutFrom_ = Trim("" & selRec!WaterBoutFrom)
    If isExist(selRec, "WaterBoutW") Then WaterBoutW_ = Trim("" & selRec!WaterBoutW)
    If isExist(selRec, "WaterBoutBack") Then WaterBoutBack_ = Trim("" & selRec!WaterBoutBack)
    If isExist(selRec, "WaterBoutTo") Then WaterBoutTo_ = Trim("" & selRec!WaterBoutTo)
    If isExist(selRec, "BoutCage") Then BoutCage_ = Trim("" & selRec!BoutCage)
    If isExist(selRec, "BoutCageFrom") Then BoutCageFrom_ = Trim("" & selRec!BoutCageFrom)
    If isExist(selRec, "BoutCageW") Then BoutCageW_ = Trim("" & selRec!BoutCageW)
    If isExist(selRec, "BoutCageBack") Then BoutCageBack_ = Trim("" & selRec!BoutCageBack)
    If isExist(selRec, "BoutCageTo") Then BoutCageTo_ = Trim("" & selRec!BoutCageTo)

End Sub

Public Sub SetDataRS_ExtCom(selRec As ADODB.Recordset)
On Error Resume Next
    id_ = Trim("" & selRec("CustomerID"))
    CustomerCode_ = Trim("" & selRec("CustomerCode"))
    custName_ = Trim("" & selRec("CompanyName"))
    provinceName_ = Trim("" & selRec("ProvinceName"))
    underSaleName_ = Trim("" & selRec("SaleStaffFName")) & " " & Trim("" & selRec("SaleStaffLName"))
End Sub
