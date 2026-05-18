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
Dim colOtherRemark_ As Collection
Dim billingInfo_ As BWGCustomerBillInfo
Dim saleDistrictID_ As String
Dim saleDistrictName_ As String
Dim saleZoneID_ As String
Dim saleZoneName_ As String
Dim custRemark_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get CustomerCode() As String
    CustomerCode = CustomerCode_
End Property
Public Property Let CustomerCode(ByVal newVal As String)
    If Trim(CustomerCode_) <> Trim(newVal) Then
        CustomerCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerName() As String
    CustomerName = custName_
End Property
Public Property Let CustomerName(ByVal newVal As String)
    If Trim(custName_) <> Trim(newVal) Then
        custName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustomerShortName() As String
    CustomerShortName = custShortName_
End Property
Public Property Let CustomerShortName(ByVal newVal As String)
    If Trim(custShortName_) <> Trim(newVal) Then
        custShortName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SaleDistrictID() As String
    SaleDistrictID = saleDistrictID_
End Property
Public Property Let SaleDistrictID(ByVal newVal As String)
    If Trim(saleDistrictID_) <> Trim(newVal) Then
        saleDistrictID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SaleDistrictName() As String
    SaleDistrictName = saleDistrictName_
End Property
Public Property Let SaleDistrictName(ByVal newVal As String)
    If Trim(saleDistrictName_) <> Trim(newVal) Then
        saleDistrictName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SaleZoneID() As String
    SaleZoneID = saleZoneID_
End Property
Public Property Let SaleZoneID(ByVal newVal As String)
    If Trim(saleZoneID_) <> Trim(newVal) Then
        saleZoneID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SaleZoneName() As String
    SaleZoneName = saleZoneName_
End Property
Public Property Let SaleZoneName(ByVal newVal As String)
    If Trim(saleZoneName_) <> Trim(newVal) Then
        saleZoneName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustRemark() As String
    CustRemark = custRemark_
End Property
Public Property Let CustRemark(ByVal newVal As String)
    If Trim(custRemark_) <> Trim(newVal) Then
        custRemark_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Set BillingInfo(newVal As BWGCustomerBillInfo)
    Set billingInfo_ = newVal
    isChanged_ = True
End Property

Public Property Get CustomerType() As String
    CustomerType = custType_
End Property
Public Property Let CustomerType(ByVal newVal As String)
    If Trim(custType_) <> Trim(newVal) Then
        custType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ParentCustID() As String
    ParentCustID = parentCustID_
End Property
Public Property Let ParentCustID(ByVal newVal As String)
    If Trim(parentCustID_) <> Trim(newVal) Then
        parentCustID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CreatedStaffID() As String
    CreatedStaffID = createdStaffID_
End Property
Public Property Let CreatedStaffID(ByVal newVal As String)
    If Trim(createdStaffID_) <> Trim(newVal) Then
        createdStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CreatedStaffName() As String
    CreatedStaffName = createdStaffName_
End Property
Public Property Let CreatedStaffName(ByVal newVal As String)
    If Trim(createdStaffName_) <> Trim(newVal) Then
        createdStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CreatedDate() As String
    CreatedDate = createdDate_
End Property
Public Property Let CreatedDate(ByVal newVal As String)
    If Trim(createdDate_) <> Trim(newVal) Then
        createdDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get UnderSaleStaffID() As String
    UnderSaleStaffID = underSaleStaffID_
End Property
Public Property Let UnderSaleStaffID(ByVal newVal As String)
    If Trim(underSaleStaffID_) <> Trim(newVal) Then
        underSaleStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get UnderSaleName() As String
    UnderSaleName = underSaleName_
End Property
Public Property Let UnderSaleName(ByVal newVal As String)
    If Trim(underSaleName_) <> Trim(newVal) Then
        underSaleName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get UnderSaleTelNo() As String
    UnderSaleTelNo = underSaleTelNo_
End Property
Public Property Let UnderSaleTelNo(ByVal newVal As String)
    If Trim(underSaleTelNo_) <> Trim(newVal) Then
        underSaleTelNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustGroupID() As String
    CustGroupID = custGroupID_
End Property
Public Property Let CustGroupID(ByVal newVal As String)
    If Trim(custGroupID_) <> Trim(newVal) Then
        custGroupID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CustGroupName() As String
    CustGroupName = custGroupName_
End Property
Public Property Let CustGroupName(ByVal newVal As String)
    If Trim(custGroupName_) <> Trim(newVal) Then
        custGroupName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ParentCustCode() As String
    ParentCustCode = parentCustCode_
End Property
Public Property Let ParentCustCode(ByVal newVal As String)
    If Trim(parentCustCode_) <> Trim(newVal) Then
        parentCustCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ParentCustName() As String
    ParentCustName = parentCustName_
End Property
Public Property Let ParentCustName(ByVal newVal As String)
    If Trim(parentCustName_) <> Trim(newVal) Then
        parentCustName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get Address() As String
    Address = custAddr_
End Property
Public Property Let Address(ByVal newVal As String)
    If Trim(custAddr_) <> Trim(newVal) Then
        custAddr_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TumbolID() As String
    TumbolID = tumbolID_
End Property
Public Property Let TumbolID(ByVal newVal As String)
    If Trim(tumbolID_) <> Trim(newVal) Then
        tumbolID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TumbolName() As String
    TumbolName = tumbolName_
End Property
Public Property Let TumbolName(ByVal newVal As String)
    If Trim(tumbolName_) <> Trim(newVal) Then
        tumbolName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AumphurID() As String
    AumphurID = aumphurID_
End Property
Public Property Let AumphurID(ByVal newVal As String)
    If Trim(aumphurID_) <> Trim(newVal) Then
        aumphurID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AumphurName() As String
    AumphurName = aumphurName_
End Property
Public Property Let AumphurName(ByVal newVal As String)
    If Trim(aumphurName_) <> Trim(newVal) Then
        aumphurName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceID() As String
    ProvinceID = provinceID_
End Property
Public Property Let ProvinceID(ByVal newVal As String)
    If Trim(provinceID_) <> Trim(newVal) Then
        provinceID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceName() As String
    ProvinceName = provinceName_
End Property
Public Property Let ProvinceName(ByVal newVal As String)
    If Trim(provinceName_) <> Trim(newVal) Then
        provinceName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ZipCode() As String
    ZipCode = zipCode_
End Property
Public Property Let ZipCode(ByVal newVal As String)
    If Trim(zipCode_) <> Trim(newVal) Then
        zipCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TelNo() As String
    TelNo = telNo_
End Property
Public Property Let TelNo(ByVal newVal As String)
    If Trim(telNo_) <> Trim(newVal) Then
        telNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FaxNo() As String
    FaxNo = faxNo_
End Property
Public Property Let FaxNo(ByVal newVal As String)
    If Trim(faxNo_) <> Trim(newVal) Then
        faxNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get EmailAddr() As String
    EmailAddr = emailAddr_
End Property
Public Property Let EmailAddr(ByVal newVal As String)
    If Trim(emailAddr_) <> Trim(newVal) Then
        emailAddr_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get Website() As String
    Website = website_
End Property
Public Property Let Website(ByVal newVal As String)
    If Trim(website_) <> Trim(newVal) Then
        website_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ContactPerson() As String
    ContactPerson = contactPerson_
End Property
Public Property Let ContactPerson(ByVal newVal As String)
    If Trim(contactPerson_) <> Trim(newVal) Then
        contactPerson_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get PositionName() As String
    PositionName = positionName_
End Property
Public Property Let PositionName(ByVal newVal As String)
    If Trim(positionName_) <> Trim(newVal) Then
        positionName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DepartmentName() As String
    DepartmentName = deptName_
End Property
Public Property Let DepartmentName(ByVal newVal As String)
    If Trim(deptName_) <> Trim(newVal) Then
        deptName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IndustryTypeID() As String
    IndustryTypeID = industryTypeID_
End Property
Public Property Let IndustryTypeID(ByVal newVal As String)
    If Trim(industryTypeID_) <> Trim(newVal) Then
        industryTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IndustryTypeName() As String
    IndustryTypeName = industryTypeName_
End Property
Public Property Let IndustryTypeName(ByVal newVal As String)
    If Trim(industryTypeName_) <> Trim(newVal) Then
        industryTypeName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ZoneID() As String 'ปรับแก้จาก Zone เป็น Sale District
    ZoneID = zoneID_
End Property
Public Property Let ZoneID(ByVal newVal As String)
    If Trim(zoneID_) <> Trim(newVal) Then
        zoneID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ZoneName() As String
    ZoneName = zoneName_
End Property
Public Property Let ZoneName(ByVal newVal As String)
    If Trim(zoneName_) <> Trim(newVal) Then
        zoneName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get RegisterNo() As String
    RegisterNo = registerNo_
End Property
Public Property Let RegisterNo(ByVal newVal As String)
    If Trim(registerNo_) <> Trim(newVal) Then
        registerNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteManageID() As String
    WasteManageID = wasteManageID_
End Property
Public Property Let WasteManageID(ByVal newVal As String)
    If Trim(wasteManageID_) <> Trim(newVal) Then
        wasteManageID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TaxNo() As String
    TaxNo = taxNo_
End Property
Public Property Let TaxNo(ByVal newVal As String)
    If Trim(taxNo_) <> Trim(newVal) Then
        taxNo_ = Trim(newVal)
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
Public Property Let isAccountOk(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
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
'Dim tmpS As New BWGSearchManager
'    If colWasteItem_ Is Nothing Then Set colWasteItem_ = tmpS.WasteDataCRSearch("CustomerID='" & id_ & "'")
'    Set WasteItems = colWasteItem_
'    Set tmpS = Nothing
End Property
Public Property Set WasteItems(ByVal newVal As Collection)
    Set colWasteItem_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryWasteItems()
    Set colWasteItem_ = Nothing
End Sub

Public Property Get SiteAddresses() As Collection
'Dim tmpS As New BWGSearchManager
'    If colSiteAddr_ Is Nothing Then Set colSiteAddr_ = tmpS.CustomerSiteSearch("CustomerID='" & id_ & "'")
'    Set SiteAddresses = colSiteAddr_
'    Set tmpS = Nothing
End Property
Public Property Set SiteAddresses(ByVal newVal As Collection)
    Set colSiteAddr_ = newVal
    isChanged_ = True
End Property
Public Sub RequerySiteAddress()
    Set colSiteAddr_ = Nothing
End Sub

Public Property Get BillAddresses() As Collection
'Dim tmpS As New BWGSearchManager
'    If colBillAddr_ Is Nothing Then Set colBillAddr_ = tmpS.CustomerBillingAddrSearch("CustomerID='" & id_ & "'")
'    Set BillAddresses = colBillAddr_
'    Set tmpS = Nothing
End Property
Public Property Set BillAddresses(ByVal newVal As Collection)
    Set colBillAddr_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryBillAddress()
    Set colBillAddr_ = Nothing
End Sub

Public Property Get Contracts() As Collection
'Dim tmpS As New BWGSearchManager
'    If colContract_ Is Nothing Then Set colContract_ = tmpS.ContractSearch("CustomerID='" & id_ & "'")
'    Set Contracts = colContract_
'    Set tmpS = Nothing
End Property
Public Property Set Contracts(ByVal newVal As Collection)
    Set colContract_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryContracts()
    Set colContract_ = Nothing
End Sub

Public Property Get CustStaffs() As Collection
'Dim tmpS As New BWGSearchManager
'    If colCustStaff_ Is Nothing Then Set colCustStaff_ = tmpS.CustomerStaffSearch("CustomerID='" & id_ & "'")
'    Set CustStaffs = colCustStaff_
'    Set tmpS = Nothing
End Property
Public Property Set CustStaffs(ByVal newVal As Collection)
    Set colCustStaff_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryCustStaff()
    Set colContract_ = Nothing
End Sub

Public Property Get OtherRemarks() As Collection
'Dim tmpS As New BWGSearchManager
'    If colOtherRemark_ Is Nothing Then Set colOtherRemark_ = tmpS.CustOrderRemarkSearch("CustomerID='" & id_ & "'")
'    Set OtherRemarks = colOtherRemark_
'    Set tmpS = Nothing
End Property
Public Property Set OtherRemarks(ByVal newVal As Collection)
    Set colOtherRemark_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryOtherRemark()
    Set colOtherRemark_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub Save()
'Dim tmpRec As ADODB.Recordset
'Dim tmpSite As BWGCustomerSite
'Dim tmpBillAddr As BWGCustomerBillAddr
'    Set tmpRec = New ADODB.Recordset
'    tmpRec.Open "Select * from tbCustomer Where CustomerID='" & id_ & "'", DBConn, adOpenDynamic, adLockOptimistic
'    If tmpRec.EOF Then
'        tmpRec.AddNew
'        tmpRec("CustomerID") = id_
'        tmpRec("CreatedByStaffID") = Trim(createdStaffID_)
'        tmpRec("CreatedDate") = Trim(createdDate_)
'        Set tmpSite = New BWGCustomerSite
'        With tmpSite
'            .ID = GetGUID
'            .Address = Trim(custAddr_)
'            .SiteName = "ที่ตั้งโรงงาน"
'            .TumbolID = Trim(tumbolID_)
'            .ZipCode = Trim(zipCode_)
'            .CustomerID = id_
'            .TelNo = Trim(telNo_)
'            .FaxNo = Trim(faxNo_)
'        End With
'        Set colSiteAddr_ = Nothing
'        Set tmpBillAddr = New BWGCustomerBillAddr
'        With tmpBillAddr
'            .ID = GetGUID
'            .BillName = "ที่ตั้งโรงงาน"
'            .CustomerID = id_
'            .FaxNo = Trim(faxNo_)
'            .TelNo = Trim(telNo_)
'            .TumbolID = Trim(tumbolID_)
'            .ZipCode = Trim(zipCode_)
'        End With
'        Set colBillAddr_ = Nothing
'    End If
'    tmpRec("CustRemark") = Trim(custRemark_)
'    tmpRec("CompanyName") = Trim(custName_)
'    tmpRec("CustShortName") = Trim(custShortName_)
'    tmpRec("Addr1") = Trim(custAddr_)
'    tmpRec("TumbolID") = Trim(tumbolID_)
'    tmpRec("ZipCode") = Trim(zipCode_)
'    tmpRec("TelNo") = Trim(telNo_)
'    tmpRec("FaxNo") = Trim(faxNo_)
'    tmpRec("WebSite") = Trim(website_)
'    tmpRec("EmailAddr") = Trim(emailAddr_)
'    tmpRec("ContactPerson") = Trim(contactPerson_)
'    tmpRec("PositionName") = Trim(positionName_)
'    tmpRec("ContactDeptNo") = Trim(deptName_)
'    tmpRec("IndustryTypeID") = Trim(industryTypeID_)
'    tmpRec("ZoneID") = Trim(zoneID_)
'    tmpRec("RegisterNo") = Trim(registerNo_)
'    tmpRec("WasteManageID") = Trim(wasteManageID_)
'    tmpRec("TaxNo") = Trim(taxNo_)
'    tmpRec("isAccountantOK") = Trim(isAccountOK_)
'    tmpRec("CustTypeFlag") = Trim(custType_)
'    tmpRec("ParentCustID") = Trim(parentCustID_)
'    tmpRec("UnderSaleStaffID") = Trim(underSaleStaffID_)
'    tmpRec("CustGroupID") = Trim(custGroupID_)
'    tmpRec("SaleDistrictID") = Trim(saleDistrictID_)
'    tmpRec.Update
'    tmpRec.Requery
'    customerCode_ = Trim("" & tmpRec("CustomerCode"))
'    If Not billingInfo_ Is Nothing Then
'        billingInfo_.Save
'    End If
'    tmpRec.Close
'    If Not tmpSite Is Nothing Then tmpSite.Save
'    If Not tmpBillAddr Is Nothing Then tmpBillAddr.Save
'    Set tmpSite = Nothing
'    Set tmpBillAddr = Nothing
'    Set tmpRec = Nothing
'    isChanged_ = False
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbCustomer Where CustomerID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("CustomerID"))
    CustomerCode_ = Trim("" & selRec("CustomerCode"))
    custName_ = Trim("" & selRec("CompanyName"))
    custShortName_ = Trim("" & selRec("CustShortName"))
    custAddr_ = Trim("" & selRec("Addr1"))
    tumbolID_ = Trim("" & selRec("TumbolID"))
    tumbolName_ = Trim("" & selRec("TumbolName"))
    aumphurID_ = Trim("" & selRec("AumphurID"))
    aumphurName_ = Trim("" & selRec("AumphurName"))
    provinceID_ = Trim("" & selRec("ProvinceID"))
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
    underSaleName_ = Trim("" & selRec("SaleStaffFName")) & " " & Trim("" & selRec("SaleStaffLName"))
    custGroupID_ = Trim("" & selRec("CustGroupID"))
    custGroupName_ = Trim("" & selRec("CustGroupName"))
    createdStaffName_ = Trim("" & selRec("CreatedStaffFName")) & " " & Trim("" & selRec("CreatedStaffLName"))
    saleDistrictID_ = Trim("" & selRec("SaleDistrictID"))
    saleDistrictName_ = Trim("" & selRec("SDName"))
    saleZoneID_ = Trim("" & selRec("SaleZoneID"))
    saleZoneName_ = Trim("" & selRec("SaleZoneName"))
    custRemark_ = Trim("" & selRec("CustRemark"))
    underSaleTelNo_ = Trim("" & selRec("SaleTelNo"))
End Sub

Public Sub SetDataRS_ExtCom(selRec As ADODB.Recordset)
On Error Resume Next
    id_ = Trim("" & selRec("CustomerID"))
    CustomerCode_ = Trim("" & selRec("CustomerCode"))
    custName_ = Trim("" & selRec("CustomerName"))
    provinceName_ = Trim("" & selRec("ProvinceName"))
    underSaleName_ = Trim("" & selRec("SaleStaffFName")) & " " & Trim("" & selRec("SaleStaffLName"))
End Sub
