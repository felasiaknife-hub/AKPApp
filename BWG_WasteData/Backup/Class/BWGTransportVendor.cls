VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGTransportVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim venderType_ As String
Dim venderCode_ As String
Dim nameTH_ As String
Dim nameEN_ As String
Dim age_ As String
Dim addr_ As String
Dim tumbolID_ As String
Dim tumbolName_ As String
Dim aumphurID_ As String
Dim aumphurName_ As String
Dim provinceID_ As String
Dim provinceName_ As String
Dim zipCode_ As String
Dim taxID_ As String
Dim telNo_ As String
Dim faxNo_ As String
Dim occupancy_ As String
Dim emailAddr_ As String
Dim website_ As String
Dim creditDay_ As Long
Dim discountPC_ As Double
Dim supplierRem_ As String
Dim isDeleted_ As String
Dim deleteDate_ As String
Dim createDate_ As String
Dim docTrans1_ As String
Dim docTrans2_ As String
Dim docTrans3_ As String
Dim docTrans4_ As String
Dim docTrans5_ As String
Dim docTrans6_ As String
Dim docTrans7_ As String
Dim transLicenseNo_ As String
Dim wTax_ As Integer
Dim workerCharge_ As Double
Dim hasVAT_ As Integer
Dim colSupplierContact_ As Collection
Dim colVenderCar_ As Collection
Dim isChanged_ As Boolean
Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get VendorType() As String
    VendorType = venderType_
End Property
Public Property Let VendorType(ByVal newVal As String)
    If Trim(venderType_) <> Trim(newVal) Then
        venderType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VendorCode() As String
    VendorCode = venderCode_
End Property
Public Property Let VendorCode(ByVal newVal As String)
    If Trim(venderCode_) <> Trim(newVal) Then
        venderCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get NameTH() As String
    NameTH = nameTH_
End Property
Public Property Let NameTH(ByVal newVal As String)
    If Trim(nameTH_) <> Trim(newVal) Then
        nameTH_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get NameEN() As String
    NameEN = nameEN_
End Property
Public Property Let NameEN(ByVal newVal As String)
    If Trim(nameEN_) <> Trim(newVal) Then
        nameEN_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get Age() As String
    Age = age_
End Property
Public Property Let Age(ByVal newVal As String)
    If Trim(age_) <> Trim(newVal) Then
        age_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get Addr() As String
    Addr = addr_
End Property
Public Property Let Addr(ByVal newVal As String)
    If Trim(addr_) <> Trim(newVal) Then
        addr_ = Trim(newVal)
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

Public Property Get TaxID() As String
    TaxID = taxID_
End Property
Public Property Let TaxID(ByVal newVal As String)
    If Trim(taxID_) <> Trim(newVal) Then
        taxID_ = Trim(newVal)
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

Public Property Get Occupancy() As String
    Occupancy = occupancy_
End Property
Public Property Let Occupancy(ByVal newVal As String)
    If Trim(occupancy_) <> Trim(newVal) Then
        occupancy_ = Trim(newVal)
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

Public Property Get CreditDay() As Long
    CreditDay = creditDay_
End Property
Public Property Let CreditDay(ByVal newVal As Long)
    If creditDay_ <> newVal Then
        creditDay_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get DiscountPC() As Double
    DiscountPC = discountPC_
End Property
Public Property Let DiscountPC(ByVal newVal As Double)
    If discountPC_ <> newVal Then
        discountPC_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get SupplierRem() As String
    SupplierRem = supplierRem_
End Property
Public Property Let SupplierRem(ByVal newVal As String)
    If Trim(supplierRem_) <> Trim(newVal) Then
        supplierRem_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isDeleted() As Boolean
    If Trim(isDeleted_) = "" Then Trim(isDeleted_) = "Y"
    If Trim(isDeleted_) = "Y" Then
        isDeleted = True
    Else
        isDeleted = False
    End If
End Property
Public Property Let isDeleted(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isDeleted_) <> Trim(tmpStr) Then
        isDeleted_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get DocTrans1() As Boolean
   ' If Trim(docTrans1_) = "" Then Trim(docTrans1_) = ""
    If Trim(docTrans1_) = "Y" Then
        DocTrans1 = True
    Else
        DocTrans1 = False
    End If
End Property
Public Property Let DocTrans1(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docTrans1_) <> Trim(tmpStr) Then
        docTrans1_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get DocTrans2() As Boolean
    'If Trim(docTrans2_) = "" Then Trim(docTrans2_) = ""
    If Trim(docTrans2_) = "Y" Then
        DocTrans2 = True
    Else
        DocTrans2 = False
    End If
End Property
Public Property Let DocTrans2(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docTrans2_) <> Trim(tmpStr) Then
        docTrans2_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get DocTrans3() As Boolean
   ' If Trim(docTrans3_) = "" Then Trim(docTrans3_) = ""
    If Trim(docTrans3_) = "Y" Then
        DocTrans3 = True
    Else
        DocTrans3 = False
    End If
End Property
Public Property Let DocTrans3(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docTrans3_) <> Trim(tmpStr) Then
        docTrans3_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get DocTrans4() As Boolean
    'If Trim(docTrans4_) = "" Then Trim(docTrans4_) = ""
    If Trim(docTrans4_) = "Y" Then
        DocTrans4 = True
    Else
        DocTrans4 = False
    End If
End Property
Public Property Let DocTrans4(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docTrans4_) <> Trim(tmpStr) Then
        docTrans4_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get DocTrans5() As Boolean
    'If Trim(docTrans5_) = "" Then Trim(docTrans5_) = ""
    If Trim(docTrans5_) = "Y" Then
        DocTrans5 = True
    Else
        DocTrans5 = False
    End If
End Property
Public Property Let DocTrans5(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docTrans5_) <> Trim(tmpStr) Then
        docTrans5_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get DocTrans6() As Boolean
   ' If Trim(docTrans6_) = "" Then Trim(docTrans6_) = ""
    If Trim(docTrans6_) = "Y" Then
        DocTrans6 = True
    Else
        DocTrans6 = False
    End If
End Property
Public Property Let DocTrans6(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docTrans6_) <> Trim(tmpStr) Then
        docTrans6_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get DocTrans7() As Boolean
   ' If Trim(docTrans7_) = "" Then Trim(docTrans7_) = ""
    If Trim(docTrans7_) = "Y" Then
        DocTrans7 = True
    Else
        DocTrans7 = False
    End If
End Property
Public Property Let DocTrans7(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(docTrans7_) <> Trim(tmpStr) Then
        docTrans7_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get TransLicenseNo() As String
    TransLicenseNo = transLicenseNo_
End Property
Public Property Let TransLicenseNo(ByVal newVal As String)
    If Trim(transLicenseNo_) <> Trim(newVal) Then
        transLicenseNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WTaxRate() As Integer
    WTaxRate = wTax_
End Property
Public Property Let WTaxRate(ByVal newVal As Integer)
    If wTax_ <> newVal Then
        wTax_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get WorkerCharge() As Double
    WorkerCharge = workerCharge_
End Property
Public Property Let WorkerCharge(ByVal newVal As Double)
    If workerCharge_ <> newVal Then
        workerCharge_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get HasVAT() As Integer
    HasVAT = hasVAT_
End Property
Public Property Let HasVAT(ByVal newVal As Integer)
    If hasVAT_ <> newVal Then
        hasVAT_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get DeletedDate() As String
    DeletedDate = deleteDate_
End Property
Public Property Let DeletedDate(ByVal newVal As String)
    If Trim(deleteDate_) <> Trim(newVal) Then
        deleteDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CreatedDate() As String
    CreatedDate = createDate_
End Property
Public Property Let CreatedDate(ByVal newVal As String)
    If Trim(createDate_) <> Trim(newVal) Then
        createDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SupplierContact() As Collection
Dim tmpS As New BWGSearchManager
    If colSupplierContact_ Is Nothing Then Set colSupplierContact_ = tmpS.SupplierContactSearch("VendorID='" & id_ & "'")
    Set SupplierContact = colSupplierContact_
End Property
Public Property Set SupplierContact(ByVal newVal As Collection)
    Set colSupplierContact_ = newVal
    isChanged_ = True
End Property
Public Sub RequerySupplierContact()
    Set colSupplierContact_ = Nothing
End Sub

Public Property Get VenderCar() As Collection
Dim tmpS As New BWGSearchManager
    If colVenderCar_ Is Nothing Then Set colVenderCar_ = tmpS.VenderCarSearch("VendorID='" & id_ & "'")
    Set VenderCar = colVenderCar_
End Property
Public Property Set VenderCar(ByVal newVal As Collection)
    Set colVenderCar_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryVenderCar()
    Set colVenderCar_ = Nothing
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
Dim x%
    
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbTransportVendor Where VendorID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("VendorID") = id_
    End If
    tmpRec("VendorType") = Trim(venderType_)
    tmpRec("NameTH") = Trim(nameTH_)
    tmpRec("NameEN") = Trim(nameEN_)
    tmpRec("TumbolID") = Trim(tumbolID_)
    tmpRec("TaxID") = Trim(taxID_)
    tmpRec("TelNo") = Trim(telNo_)
    tmpRec("FaxNo") = Trim(faxNo_)
    tmpRec("Occupancy") = Trim(occupancy_)
    tmpRec("EmailAddr") = Trim(emailAddr_)
    tmpRec("Website") = Trim(website_)
    tmpRec("CreditDay") = Trim(creditDay_)
    tmpRec("DiscountPC") = Trim(discountPC_)
    tmpRec("SupplierRem") = Trim(supplierRem_)
    tmpRec("isDeleted") = Trim(isDeleted_)
    tmpRec("DeletedDate") = Trim(deleteDate_)
    tmpRec("CreatedDate") = Trim(createDate_)
    tmpRec("Age") = Trim(age_)
    tmpRec("Addr") = Trim(addr_)
    tmpRec("ZipCode") = Trim(zipCode_)
    tmpRec("DocTrans1") = Trim(docTrans1_)
    tmpRec("DocTrans2") = Trim(docTrans2_)
    tmpRec("DocTrans3") = Trim(docTrans3_)
    tmpRec("DocTrans4") = Trim(docTrans4_)
    tmpRec("DocTrans5") = Trim(docTrans5_)
    tmpRec("DocTrans6") = Trim(docTrans6_)
    tmpRec("DocTrans7") = Trim(docTrans7_)
    tmpRec("TransLicenseNo") = Trim(transLicenseNo_)
    tmpRec("WTaxRate") = wTax_
    tmpRec("WorkerCharge") = workerCharge_
    tmpRec("HasVAT") = hasVAT_
    tmpRec.Update
    tmpRec.Requery
    venderCode_ = Trim("" & tmpRec("VendorCode"))
    tmpRec.Close
        
    With DBManager.GetDBConnection
        .Execute "Delete tbVendorCar Where VendorID='" & id_ & "' And VendorTruckID not in (Select VendorTruckID from tbJobDataTimeTable)"
        .Execute "Delete tbSupplierContact Where VendorID='" & id_ & "'"
        If Not colVenderCar_ Is Nothing Then
            For x = 1 To colVenderCar_.Count
                colVenderCar_(x).Save
            Next
        End If
        If Not colSupplierContact_ Is Nothing Then
            For x = 1 To colSupplierContact_.Count
                colSupplierContact_(x).Save
            Next
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
Dim tmpRec As New ADODB.Recordset
    With DBManager.GetDBConnection
        Set tmpRec = GetRS("Select * from tbJobDataTimeTable Where VendorTruckID in (Select VendorTruckID from tbVendorCar Where VendorID='" & id_ & "')", adOpenForwardOnly, adLockReadOnly)
'        tmpRec.Open "Select * from tbJobDataTimeTable Where VendorTruckID in (Select VendorTruckID from tbVendorCar Where VendorID='" & id_ & "')", DBManager.GetDBConnection, adOpenKeyset
        If tmpRec.EOF Then
            .Execute "Delete tbVendorCar Where VendorID='" & id_ & "'"
            .Execute "Delete tbSupplierContact Where VendorID='" & id_ & "'"
            .Execute "Delete tbTransportVendor Where VendorID='" & id_ & "'"
            Delete = ""
        Else
            Delete = "ไม่สามารถลบข้อมูล Vendor นี้ได้เพราะมีการใช้งานรถของ Vendor นี้แล้ว"
        End If
        tmpRec.Close
        Set tmpRec = Nothing
    End With
    Exit Function
delErr:
    Delete = Err.Description
    If tmpRec.State = adStateOpen Then tmpRec.Close
    Set tmpRec = Nothing
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("VendorID")
    venderType_ = Trim("" & selRec("VendorType"))
    venderCode_ = Trim("" & selRec("VendorCode"))
    nameTH_ = Trim("" & selRec("NameTH"))
    nameEN_ = Trim("" & selRec("NameEN"))
    tumbolID_ = Trim("" & selRec("TumbolID"))
    tumbolName_ = Trim("" & selRec("TumbolName"))
    aumphurID_ = Trim("" & selRec("AumphurID"))
    aumphurName_ = Trim("" & selRec("AumphurName"))
    provinceID_ = Trim("" & selRec("ProvinceID"))
    provinceName_ = Trim("" & selRec("ProvinceName"))
    taxID_ = Trim("" & selRec("TaxID"))
    telNo_ = Trim("" & selRec("TelNo"))
    faxNo_ = Trim("" & selRec("FaxNo"))
    occupancy_ = Trim("" & selRec("Occupancy"))
    emailAddr_ = Trim("" & selRec("EmailAddr"))
    website_ = Trim("" & selRec("Website"))
    creditDay_ = Trim("" & selRec("CreditDay"))
    discountPC_ = Trim("" & selRec("DiscountPC"))
    supplierRem_ = Trim("" & selRec("SupplierRem"))
    isDeleted_ = Trim("" & selRec("isDeleted"))
    deleteDate_ = Trim("" & selRec("DeletedDate"))
    createDate_ = Trim("" & selRec("CreatedDate"))
    age_ = Trim("" & selRec("Age"))
    addr_ = Trim("" & selRec("Addr"))
    zipCode_ = Trim("" & selRec("ZipCode"))
    docTrans1_ = Trim("" & selRec("DocTrans1"))
    docTrans2_ = Trim("" & selRec("DocTrans2"))
    docTrans3_ = Trim("" & selRec("DocTrans3"))
    docTrans4_ = Trim("" & selRec("DocTrans4"))
    docTrans5_ = Trim("" & selRec("DocTrans5"))
    docTrans6_ = Trim("" & selRec("DocTrans6"))
    docTrans7_ = Trim("" & selRec("DocTrans7"))
    transLicenseNo_ = Trim("" & selRec("TransLicenseNo"))
    If IsNull(selRec("WTaxRate")) = False Then wTax_ = selRec("WTaxRate")
    If IsNull(selRec("WorkerCharge")) = False Then workerCharge_ = selRec("WorkerCharge")
    If IsNull(selRec("HasVAT")) = False Then hasVAT_ = selRec("HasVAT")
End Sub

