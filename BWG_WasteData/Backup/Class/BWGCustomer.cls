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
Dim customerCode_ As String
Dim CustomerCodeDemo_ As String
Dim custName_ As String
Dim custShortName_ As String
Dim custAddr_ As String
Dim tumbolID_ As String
Dim tumbolName_ As String
Dim aumphurID_ As String
Dim aumphurName_ As String
Dim provinceID_ As String
Dim provinceName_ As String
Dim zipCode_ As String
Dim telNo_ As String
Dim faxNo_ As String
Dim website_ As String
Dim emailAddr_ As String
Dim contactPerson_ As String
Dim positionName_ As String
Dim deptName_ As String
Dim industryTypeID_ As String
Dim industryTypeName_ As String
Dim zoneID_ As String
Dim zoneName_ As String
Dim registerNo_ As String
Dim wasteManageID_ As String
Dim industryTypeIDNew_ As String
Dim TypeDesc_ As String
Dim TypeCode_ As String
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
Dim colCustStaff_ As Collection
Dim colCustQuo_ As Collection
Dim colOtherRemark_ As Collection
Dim billingInfo_ As BWGCustomerBillInfo
Dim saleDistrictID_ As String
Dim saleDistrictName_ As String
Dim saleDistrictName2_ As String
Dim saleZoneID_ As String
Dim saleZoneName_ As String
Dim custRemark_ As String
Dim isChanged_ As Boolean
Dim Status_ As String
Dim StDate_ As String
Dim Refcode_ As String
Dim isApprove_ As String
Dim ApproveDate_ As String
Dim isNew_ As String
Dim isReject_ As String
Dim RejectDate_ As String
Dim RejectDesc_ As String
Dim RejectDesc2_ As String
Dim isRequest_ As String
Dim RequestDate_ As String
Dim isApproveName_ As String
Dim Sale2ID_ As String
Dim Sale2Name_ As String
Dim WasteSK2Exp_ As String
Dim ADesc_ As String
Dim Digit14_ As String
Dim IsVIPCR_ As String
Dim EstateID_ As Integer
Dim CustdeptName_ As String
Dim CRStaffID_ As String
Dim CRStaffName_ As String
Dim CompanyPolicy_ As String

Public Property Get CRStaffName() As String
    CRStaffName = CRStaffName_
End Property
Public Property Let CRStaffName(ByVal newval As String)
    If Trim(CRStaffName_) <> Trim(newval) Then
        CRStaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CompanyPolicy() As String
    CompanyPolicy = CompanyPolicy_
End Property
Public Property Let CompanyPolicy(ByVal newval As String)
    If Trim(CompanyPolicy_) <> Trim(newval) Then
        CompanyPolicy_ = Trim(newval)
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

Public Property Get EstateID() As Double
    EstateID = EstateID_
End Property
Public Property Let EstateID(ByVal newval As Double)
    If EstateID_ <> newval Then
        EstateID_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
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
Public Property Get CustomerCode() As String
    CustomerCode = customerCode_
End Property
Public Property Let CustomerCode(ByVal newval As String)
    If Trim(customerCode_) <> Trim(newval) Then
        customerCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CustomerCodeDemo() As String
    CustomerCodeDemo = CustomerCodeDemo_
End Property
Public Property Let CustomerCodeDemo(ByVal newval As String)
    If Trim(CustomerCodeDemo_) <> Trim(newval) Then
        CustomerCodeDemo_ = Trim(newval)
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
Public Property Get SaleDistrictName2() As String
    SaleDistrictName2 = saleDistrictName2_
End Property
Public Property Let SaleDistrictName2(ByVal newval As String)
    If Trim(saleDistrictName2_) <> Trim(newval) Then
        saleDistrictName2_ = Trim(newval)
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

Public Property Get CustdeptName() As String
    CustdeptName = CustdeptName_
End Property
Public Property Let CustdeptName(ByVal newval As String)
    If Trim(CustdeptName_) <> Trim(newval) Then
        CustdeptName_ = Trim(newval)
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
    isChanged_ = True
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

Public Property Get WasteSK2Exp() As String
    WasteSK2Exp = WasteSK2Exp_
End Property
Public Property Let WasteSK2Exp(ByVal newval As String)
    If Trim(WasteSK2Exp_) <> Trim(newval) Then
        WasteSK2Exp_ = Trim(newval)
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
    Address = custAddr_
End Property
Public Property Let Address(ByVal newval As String)
    If Trim(custAddr_) <> Trim(newval) Then
        custAddr_ = Trim(newval)
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
Public Property Get ADesc() As String
    ADesc = ADesc_
End Property
Public Property Let ADesc(ByVal newval As String)
    If Trim(ADesc_) <> Trim(newval) Then
        ADesc_ = Trim(newval)
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

Public Property Get ZipCode() As String
    ZipCode = zipCode_
End Property
Public Property Let ZipCode(ByVal newval As String)
    If Trim(zipCode_) <> Trim(newval) Then
        zipCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TelNo() As String
    TelNo = telNo_
End Property
Public Property Let TelNo(ByVal newval As String)
    If Trim(telNo_) <> Trim(newval) Then
        telNo_ = Trim(newval)
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
        industryTypeIDNew_ = Trim(newval)
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
Public Property Get TypeCode() As String
    TypeCode = TypeCode_
End Property
Public Property Let TypeCode(ByVal newval As String)
    If Trim(TypeCode_) <> Trim(newval) Then
        TypeCode_ = Trim(newval)
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
    RegisterNo = registerNo_
End Property
Public Property Let RegisterNo(ByVal newval As String)
    If Trim(registerNo_) <> Trim(newval) Then
        registerNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Digit14() As String
    Digit14 = Digit14_
End Property
Public Property Let Digit14(ByVal newval As String)
    If Trim(Digit14_) <> Trim(newval) Then
        Digit14_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteManageID() As String
    WasteManageID = wasteManageID_
End Property
Public Property Let WasteManageID(ByVal newval As String)
    If Trim(wasteManageID_) <> Trim(newval) Then
        wasteManageID_ = Trim(newval)
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
    isChanged_ = True
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

Public Property Get CustStaffs() As Collection
Dim tmpS As New BWGSearchManager
    'If colCustStaff_ Is Nothing Then Set colCustStaff_ = tmpS.CustomerStaffSearch("CustomerID='" & id_ & "'")
   Set colCustStaff_ = tmpS.CustomerStaffSearch("CustomerID='" & id_ & "'")
    Set CustStaffs = colCustStaff_
    Set tmpS = Nothing
End Property
Public Property Get CustQuotations() As Collection
Dim tmpS As New BWGSearchManager
    If colCustQuo_ Is Nothing Then Set colCustQuo_ = tmpS.QuotationHeaderSearch("CustomerID='" & id_ & "'")
    'Set colCustQuo_ = tmpS.QuotationHeaderSearch("CustomerID='" & id_ & "'")
    Set CustQuotations = colCustQuo_
    Set tmpS = Nothing
End Property

Public Property Set CustStaffs(ByVal newval As Collection)
    Set colCustStaff_ = newval
    isChanged_ = True
End Property
Public Sub RequeryCustStaff()
    Set colContract_ = Nothing
End Sub

Public Property Set CustQuotations(ByVal newval As Collection)
    Set colCustQuo_ = newval
    isChanged_ = True
End Property
Public Sub RequeryCustQuotations()
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

Public Property Get isApprove() As Boolean
    If Trim(isApprove_) = "Y" Then
        isApprove = True
    Else
        isApprove = False
    End If
End Property
Public Property Let isApprove(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isApprove_) <> Trim(tmpStr) Then
        isApprove_ = Trim(tmpStr)
    End If
End Property

Public Property Get ApproveDate() As String
    ApproveDate = ApproveDate_
End Property
Public Property Let ApproveDate(ByVal newval As String)
    If Trim(ApproveDate_) <> Trim(newval) Then
        ApproveDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isReject() As Boolean
    If Trim(isReject_) = "Y" Then
        isReject = True
    Else
        isReject = False
    End If
End Property
Public Property Let isReject(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isReject_) <> Trim(tmpStr) Then
        isReject_ = Trim(tmpStr)
    End If
End Property

Public Property Get RejectDate() As String
    RejectDate = RejectDate_
End Property
Public Property Let RejectDate(ByVal newval As String)
    If Trim(RejectDate_) <> Trim(newval) Then
        RejectDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get RejectDesc() As String
    RejectDesc = RejectDesc_
End Property
Public Property Let RejectDesc(ByVal newval As String)
    If Trim(RejectDesc_) <> Trim(newval) Then
        RejectDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get RejectDesc2() As String
    RejectDesc2 = RejectDesc2_
End Property
Public Property Let RejectDesc2(ByVal newval As String)
    If Trim(RejectDesc2_) <> Trim(newval) Then
        RejectDesc2_ = Trim(newval)
        isChanged_ = True
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

Public Property Get isRequest() As Boolean
    If Trim(isRequest_) = "Y" Then
        isRequest = True
    Else
        isRequest = False
    End If
End Property
Public Property Let isRequest(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isRequest_) <> Trim(tmpStr) Then
        isRequest_ = Trim(tmpStr)
    End If
End Property
Public Property Get RequestDate() As String
    RequestDate = RequestDate_
End Property
Public Property Let RequestDate(ByVal newval As String)
    If Trim(RequestDate_) <> Trim(newval) Then
        RequestDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get isApproveName() As String
 isApproveName = isApproveName_
End Property
Public Property Let isApproveName(ByVal newval As String)
isApproveName_ = isApproveName
End Property
Public Property Get IsVIPCR() As String
 IsVIPCR = IsVIPCR_
End Property
Public Property Let IsVIPCR(ByVal newval As String)
IsVIPCR_ = IsVIPCR
End Property

Public Property Get Status() As String
Status = Status_
End Property
Public Property Let Status(ByVal newval As String)
Status_ = Status
End Property
Public Property Get StDate() As String
StDate = StDate_
End Property
Public Property Let StDate(ByVal newval As String)
StDate_ = StDate
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

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim tmpSite As BWGCustomerSite
Dim tmpBillAddr As BWGCustomerBillAddr
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbCustomer Where CustomerID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("CustomerID") = id_
        tmpRec("CreatedByStaffID") = Trim(createdStaffID_)
        tmpRec("CreatedDate") = Trim(createdDate_)
         tmpRec("Status") = "D"
        If isNew_ = "N" Then
            Set tmpSite = New BWGCustomerSite
            With tmpSite
                .ID = GetGUID
                .Address = Trim(custAddr_)
                .SiteName = "ที่ตั้งโรงงาน"
                .TumbolID = Trim(tumbolID_)
                .ZipCode = Trim(zipCode_)
                .CustomerID = id_
                .TelNo = Trim(telNo_)
                .FaxNo = Trim(faxNo_)
            End With
            Set colSiteAddr_ = Nothing
            Set tmpBillAddr = New BWGCustomerBillAddr
            With tmpBillAddr
                .ID = GetGUID
                .BillName = "ที่ตั้งโรงงาน"
                .CustomerID = id_
                .FaxNo = Trim(faxNo_)
                .TelNo = Trim(telNo_)
                .TumbolID = Trim(tumbolID_)
                .ZipCode = Trim(zipCode_)
            End With
            Set colBillAddr_ = Nothing
        End If
    End If
    tmpRec("CustRemark") = Trim(custRemark_)
    tmpRec("CompanyName") = Trim(custName_)
    tmpRec("CustShortName") = Trim(custShortName_)
    tmpRec("Addr1") = Trim(custAddr_)
    tmpRec("TumbolID") = Trim(tumbolID_)
    tmpRec("ZipCode") = Trim(zipCode_)
    tmpRec("TelNo") = Trim(telNo_)
    tmpRec("FaxNo") = Trim(faxNo_)
    tmpRec("WebSite") = Trim(website_)
    tmpRec("EmailAddr") = Trim(emailAddr_)
    tmpRec("ContactPerson") = Trim(contactPerson_)
    tmpRec("PositionName") = Trim(positionName_)
    tmpRec("ContactDeptNo") = Trim(deptName_)
    tmpRec("IndustryTypeID") = Trim(industryTypeID_)
    tmpRec("ZoneID") = Trim(zoneID_)
    tmpRec("RegisterNo") = Trim(registerNo_)
    tmpRec("14Digit") = Trim(Digit14_)
    tmpRec("IndustryTypeIDNew") = Trim(industryTypeIDNew_)
    tmpRec("WasteManageID") = Trim(wasteManageID_)
    tmpRec("TaxNo") = Trim(taxNo_)
    tmpRec("isAccountantOK") = Trim(isAccountOK_)
    tmpRec("CustTypeFlag") = Trim(custType_)
    tmpRec("ParentCustID") = Trim(parentCustID_)
    tmpRec("UnderSaleStaffID") = Trim(underSaleStaffID_)
    tmpRec("CustGroupID") = Trim(custGroupID_)
    tmpRec("SaleDistrictID") = Trim(saleDistrictID_)
    tmpRec("RefCode") = Trim(Refcode_)
    tmpRec("WasteSK2Exp") = Trim(WasteSK2Exp_)
    tmpRec("CustomerCodeDemo") = Trim(CustomerCodeDemo_)
    tmpRec("CompanyPolicy") = Trim(CompanyPolicy_)
    tmpRec!isNew = isNew_
    If isNew_ = "Y" Then tmpRec!isRequest = "N"
    tmpRec("Sale2ID") = Trim(Sale2ID_)
    tmpRec("Sale2Name") = Sale2Name_
    If EstateID_ <> "0" Then tmpRec("EstateID") = EstateID_
    tmpRec("CustDept") = CustdeptName_
    tmpRec("CRStaffID") = CRStaffID_
    tmpRec.Update
    tmpRec.Requery
    customerCode_ = Trim("" & tmpRec("CustomerCode"))
    Set tmpSite = Nothing
    Set tmpBillAddr = Nothing
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    If isNew_ = "Y" Then
        DBManager.GetDBConnection.Execute "Delete tbCustomer  Where CustomerID='" & id_ & "'"
    End If
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
Dim tmpRec As ADODB.Recordset
Dim IsDemo As String
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbquotationheader Where CustomerID='" & Trim("" & selRec("CustomerID")) & "' and isApproved='Y' ", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        IsDemo = "Y"
    Else
        IsDemo = "N"
    End If
    Set tmpRec = Nothing

    id_ = Trim("" & selRec("CustomerID"))
    If IsDemo = "N" Then
        customerCode_ = Trim("" & selRec("CustomerCode"))
    Else
        If Trim("" & selRec("CustomerCodeDemo")) <> "" Then
            customerCode_ = Trim("" & selRec("CustomerCodeDemo"))
        Else
            customerCode_ = Trim("" & selRec("CustomerCode"))
        End If
    End If
    custName_ = Trim("" & selRec("CompanyName"))
    custShortName_ = Trim("" & selRec("CustShortName"))
    custAddr_ = Trim("" & selRec("Addr1"))
    tumbolID_ = Trim("" & selRec("TumbolID"))
    tumbolName_ = Trim("" & selRec("TumbolName"))
    aumphurID_ = Trim("" & selRec("AumphurID"))
    aumphurName_ = Trim("" & selRec("AumphurName"))
    provinceID_ = Trim("" & selRec("ProvinceID"))
    WasteSK2Exp_ = IIf(IsNull(selRec!WasteSK2Exp), "", selRec!WasteSK2Exp)
    provinceName_ = Trim("" & selRec("ProvinceName"))
    zipCode_ = Trim("" & selRec("ZipCode"))
    telNo_ = Trim("" & selRec("TelNo"))
    faxNo_ = Trim("" & selRec("FaxNo"))
    website_ = Trim("" & selRec("WebSite"))
    emailAddr_ = Trim("" & selRec("EmailAddr"))
    contactPerson_ = Trim("" & selRec("ContactPerson"))
    positionName_ = Trim("" & selRec("PositionName"))
    deptName_ = Trim("" & selRec("ContactDeptNo"))
    industryTypeID_ = Trim("" & selRec("IndustryTypeID"))
    industryTypeName_ = Trim("" & selRec("TypeName"))
    zoneID_ = Trim("" & selRec("ZoneID"))
    zoneName_ = Trim("" & selRec("ZoneName"))
    registerNo_ = Trim("" & selRec("RegisterNo"))
    Digit14_ = Trim("" & selRec("14Digit"))
    wasteManageID_ = Trim("" & selRec("WasteManageID"))
    taxNo_ = Trim("" & selRec("TaxNo"))
    isAccountOK_ = Trim("" & selRec("isAccountantOK"))
    custType_ = Trim("" & selRec("CustTypeFlag"))
    parentCustID_ = Trim("" & selRec("ParentCustID"))
    parentCustCode_ = Trim("" & selRec("ParentCustCode"))
    parentCustName_ = Trim("" & selRec("ParentCustName"))
    createdStaffID_ = Trim("" & selRec("CreatedByStaffID"))
    createdDate_ = Trim("" & selRec("CreatedDate"))
    underSaleStaffID_ = Trim("" & selRec("UnderSaleStaffID"))
    underSaleName_ = Trim("" & selRec("SaleStaffFName")) & " " & Trim("" & selRec("SaleStaffLName")) & " " & Trim("" & selRec("StaffCode"))
    custGroupID_ = Trim("" & selRec("CustGroupID"))
    custGroupName_ = Trim("" & selRec("CustGroupName"))
    createdStaffName_ = Trim("" & selRec("CreatedStaffFName")) & " " & Trim("" & selRec("CreatedStaffLName"))
    saleDistrictID_ = Trim("" & selRec("SaleDistrictID"))
    saleDistrictName_ = Trim("" & selRec("SDName"))
    saleDistrictName2_ = Trim("" & selRec("SDName2"))
    saleZoneID_ = Trim("" & selRec("SaleZoneID"))
    saleZoneName_ = Trim("" & selRec("SaleZoneName"))
    custRemark_ = Trim("" & selRec("CustRemark"))
    Refcode_ = Trim("" & selRec("RefCode"))
    underSaleTelNo_ = Trim("" & selRec("SaleTelNo"))
    Status_ = Trim("" & selRec("Status"))
    isApprove_ = IIf(IsNull(selRec!isApprove), "N", selRec!isApprove)
    ApproveDate_ = IIf(IsNull(selRec!ApproveDate), "N", selRec!ApproveDate)
    isNew_ = IIf(IsNull(selRec!isNew), "N", selRec!isNew)
    isRequest_ = IIf(IsNull(selRec!isRequest), "N", selRec!isRequest)
    RequestDate_ = IIf(IsNull(selRec!RequestDate), "N", selRec!RequestDate)
    Sale2ID_ = IIf(IsNull(selRec!Sale2ID), "", selRec!Sale2ID)
    Sale2Name_ = IIf(IsNull(selRec!Sale2Name), "", selRec!Sale2Name)
    industryTypeIDNew_ = IIf(IsNull(selRec!IndustryTypeIDNew), "", selRec!IndustryTypeIDNew)
    TypeDesc_ = IIf(IsNull(selRec!TypeDesc), "", selRec!TypeDesc)
    TypeCode_ = IIf(IsNull(selRec!TypeCode), "", selRec!TypeCode)
    If IsNull(selRec!RejectDesc) = False Then
        If Trim(selRec!RejectDesc) <> "" Then
            isReject_ = "Y"
        Else
            isReject_ = "N"
        End If
    Else
        isReject_ = "N"
    End If
    RejectDesc_ = IIf(IsNull(selRec!RejectDesc), "", selRec!RejectDesc)
    RejectDesc2_ = IIf(IsNull(selRec!RejectDesc2), "", selRec!RejectDesc2)
    IsVIPCR_ = Trim("" & selRec("IsVIPCR"))
    If isExist(selRec, "EstateID") Then
        EstateID_ = IIf(IsNull(selRec!EstateID), "0", selRec!EstateID)
    End If
    If isExist(selRec, "CustDept") Then
        CustdeptName_ = IIf(IsNull(selRec!CustDept), "0", selRec!CustDept)
    End If
    If isExist(selRec, "CRStaffID") Then CRStaffID_ = Trim("" & selRec!CRStaffID)
    If isExist(selRec, "CRStaffName") Then CRStaffName_ = Trim("" & selRec!CRStaffName)
    If isExist(selRec, "CompanyPolicy") Then CompanyPolicy_ = Trim("" & selRec!CompanyPolicy)

End Sub
Public Sub UpdateSK2Exp()
Dim tSql As String
Dim RS As ADODB.Recordset
    tSql = "SELECT * FROM tbCustomer" & vbCrLf
    tSql = tSql & "WHERE CustomerID = '" & id_ & "'"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not RS.EOF Then
        tSql = "Update tbCustomer Set WasteSK2Exp='" & WasteSK2Exp_ & "' Where CustomerID='" & id_ & "'"
        DBManager.GetDBConnection.Execute tSql
    Else: Exit Sub
    End If
    Set RS = Nothing
End Sub
Public Sub UpdateIndustryType()
Dim tSql As String
Dim RS As ADODB.Recordset
    tSql = "SELECT * FROM tbCustomer" & vbCrLf
    tSql = tSql & "WHERE CustomerID = '" & id_ & "'"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
    If Not RS.EOF Then
        tSql = "Update tbCustomer Set IndustryTypeIDNew='" & industryTypeIDNew_ & "' Where CustomerID='" & id_ & "'"
        DBManager.GetDBConnection.Execute tSql
    Else: Exit Sub
    End If
    Set RS = Nothing
End Sub

Public Sub CRApprove()
Dim tSql As String
Dim RS As ADODB.Recordset
    tSql = "SELECT * FROM tbCustomer" & vbCrLf
    tSql = tSql & " WHERE CustomerID = '" & id_ & "'"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)
      Dim Approvename As String
      Approvename = "" & CurrentUser.StaffFName & " " & "   " & " " & CurrentUser.StaffLName & ""
    If Not RS.EOF Then
        DBManager.GetDBConnection.Execute "Update tbCustomer Set isNew='N', isApprove = 'Y',ApproveDate = '" & TodayDate & "' , isApproveName ='" & Approvename & "',Status = 'N'  Where CustomerID='" & id_ & "'"
        If Trim(customerCode_) = "" Or Trim(customerCode_) = Null Or Trim(CustomerCode) = "NULL" Then
            DBManager.GetDBConnection.Execute "Exec sp_CreateCustomerCode '" & id_ & "'"
        End If
    Else
        tSql = "Update tbCustomer Set isNew='N', isApprove = 'Y', ApproveDate = '" & TodayDate & "' , isApproveName ='" & Approvename & "',Status = 'N'   Where CustomerID='" & id_ & "'"
        DBManager.GetDBConnection.Execute tSql
        If Trim(customerCode_) = "" Or Trim(customerCode_) = Null Or Trim(CustomerCode) = "NULL" Then
            tSql = "Exec sp_CreateCustomerCode '" & id_ & "' "
            DBManager.GetDBConnection.Execute tSql
        End If
    End If
    Set RS = Nothing
    isApprove_ = "Y"
    ApproveDate_ = TodayDate
End Sub

Public Sub Reject(Optional ByVal RejectDesc As String = "ไม่ระบุรายละเอียด")
Dim tSql As String
Dim RS As ADODB.Recordset
    tSql = "Update tbCustomer Set RejectDesc = 'Y" & Trim(RejectDesc) & "', isRequest = 'N', RequestDate = '" & TodayDate & "', isNew='Y', isApprove = 'N',ApproveDate = '',Status = 'J'  Where CustomerID='" & id_ & "'"
    DBManager.GetDBConnection.Execute tSql
    isReject_ = "Y"
    RejectDate_ = TodayDate
End Sub

Public Sub Request()
Dim tSql As String
Dim RS As ADODB.Recordset
     tSql = "Update tbCustomer Set isRequest = 'Y', RequestDate = '" & TodayDate & "', isNew='Y', isApprove = 'N',ApproveDate = '', RejectDesc = '',Status = 'R'  Where CustomerID='" & id_ & "'"
    DBManager.GetDBConnection.Execute tSql
    isRequest_ = "Y"
    RequestDate = TodayDate
End Sub

