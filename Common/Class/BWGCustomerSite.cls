VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCustomerSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim CustomerID_ As String
Dim siteName_ As String
Dim addr1_ As String
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
Dim ContactName_ As String
Dim isManifestAddr_ As String
Dim isChanged_ As Boolean
Dim LengthNum_ As String
Dim OilNum_ As String
Dim RateSingle_ As String
Dim RateTrailer_ As String
Dim Rate15Q_ As String
Dim Rate24Q_ As String
Dim RatePickup_ As String
Dim Rate10_ As String
Dim RateTrail_ As String
Dim RunDay_ As String
Dim RunHour_ As String
Public Property Get RateSingle() As String
    RateSingle = RateSingle_
End Property
Public Property Let RateSingle(ByVal newval As String)
    If Trim(RateSingle_) <> Trim(newval) Then
        RateSingle_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get RunDay() As String
    RunDay = RunDay_
End Property
Public Property Let RunDay(ByVal newval As String)
    If Trim(RunDay_) <> Trim(newval) Then
        RunDay_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get RunHour() As String
    RunHour = RunHour_
End Property
Public Property Let RunHour(ByVal newval As String)
    If Trim(RunHour_) <> Trim(newval) Then
        RunHour_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Rate10() As String
    Rate10 = Rate10_
End Property
Public Property Let Rate10(ByVal newval As String)
    If Trim(Rate10_) <> Trim(newval) Then
        Rate10_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get RateTrail() As String
    RateTrail = RateTrail_
End Property
Public Property Let RateTrail(ByVal newval As String)
    If Trim(RateTrail_) <> Trim(newval) Then
        RateTrail_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get RatePickup() As String
    RatePickup = RatePickup_
End Property
Public Property Let RatePickup(ByVal newval As String)
       If Trim(RatePickup_) <> Trim(newval) Then
        RatePickup_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Rate24Q() As String
    Rate24Q = Rate24Q_
End Property
Public Property Let Rate24Q(ByVal newval As String)
       If Trim(Rate24Q_) <> Trim(newval) Then
        Rate24Q_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get RateTrailer() As String
    RateTrailer = RateTrailer_
End Property
Public Property Let RateTrailer(ByVal newval As String)
       If Trim(RateTrailer_) <> Trim(newval) Then
        RateTrailer_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Rate15Q() As String
    Rate15Q = Rate15Q_
End Property
Public Property Let Rate15Q(ByVal newval As String)
       If Trim(Rate15Q_) <> Trim(newval) Then
        Rate15Q_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property
Public Property Get OilNum() As String
    OilNum = OilNum_
End Property
Public Property Let OilNum(ByVal newval As String)
    If Trim(OilNum_) <> Trim(newval) Then
        OilNum_ = Trim(newval)
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
Public Property Get SiteName() As String
    SiteName = siteName_
End Property
Public Property Let SiteName(ByVal newval As String)
    If Trim(siteName_) <> Trim(newval) Then
        siteName_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get Address() As String
    Address = addr1_
End Property
Public Property Let Address(ByVal newval As String)
    If Trim(addr1_) <> Trim(newval) Then
        addr1_ = Trim(newval)
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
Public Property Get ZipCode() As String
    ZipCode = zipCode_
End Property
Public Property Let ZipCode(ByVal newval As String)
    If Trim(zipCode_) <> Trim(newval) Then
        zipCode_ = Trim(newval)
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
Public Property Get ContactName() As String
    ContactName = ContactName_
End Property
Public Property Let ContactName(ByVal newval As String)
    If Trim(ContactName_) <> Trim(newval) Then
        ContactName_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get isManifestAddr() As Boolean
    If Trim(isManifestAddr_) = "Y" Then
        isManifestAddr = True
    Else
        isManifestAddr = False
    End If
End Property
Public Property Let isManifestAddr(ByVal newval As Boolean)
Dim tmpStr As String
    If newval Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isManifestAddr_) <> Trim(tmpStr) Then
        isManifestAddr_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property
Public Property Get LengthNum() As String
    LengthNum = LengthNum_
End Property
Public Property Let LengthNum(ByVal newval As String)
       If Trim(LengthNum_) <> Trim(newval) Then
        LengthNum_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbCustomerSite WHERE CustSiteID = '" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("CustSiteID") = id_
    End If
    tmpRec("CustomerID") = Trim(CustomerID_)
    tmpRec("SiteName") = Trim(siteName_)
    tmpRec("Addr1") = Trim(addr1_)
    tmpRec("TumbolID") = Trim(tumbolID_)
    tmpRec("ZipCode") = Trim(zipCode_)
    tmpRec("Country") = Trim(Country_)
    tmpRec("TelNo") = Trim(TelNo_)
    tmpRec("FaxNo") = Trim(faxNo_)
    tmpRec("ContactName") = Trim(ContactName_)
    tmpRec("isManifestAddr") = Trim(isManifestAddr_)
    If LengthNum_ <> "" Then
    tmpRec("LengthNum") = Trim(LengthNum_)
    End If
    tmpRec("LastUser") = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec("LastUpdate") = GetSystemDateTime.SystemDateTime
    If OilNum_ <> "" Then
    tmpRec("OilNum") = Trim(OilNum_)
    End If
    If RateSingle_ <> "" Then
    tmpRec("RateSingle") = Trim(RateSingle_)
    End If
    If RateTrailer_ <> "" Then
    tmpRec("RateTrailer") = Trim(RateTrailer_)
    End If
    If Rate15Q_ <> "" Then
    tmpRec("Rate15Q") = Trim(Rate15Q_)
    End If
    If Rate24Q_ <> "" Then
    tmpRec("Rate24Q") = Trim(Rate24Q_)
    End If
    If RatePickup_ <> "" Then
    tmpRec("RatePickup") = Trim(RatePickup_)
    End If
    If Rate10_ <> "" Then
    tmpRec("Rate10") = Trim(Rate10_)
    End If
    If RateTrail_ <> "" Then
    tmpRec("RateTrail") = Trim(RateTrail_)
    End If
    If RunHour_ <> "" Then
    tmpRec("RunHour") = Trim(RunHour_)
    End If
    If RunDay_ <> "" Then
    tmpRec("RunDay") = Trim(RunDay_)
    End If
  
    tmpRec.Update
'    tmpRec.Close
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
    DBConnExc "Delete tbCustomerSite WHERE CustSiteID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("CustSiteID"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    siteName_ = Trim("" & selRec("SiteName"))
    addr1_ = Trim("" & selRec("Addr1"))
    tumbolID_ = Trim("" & selRec("TumbolID"))
    tumbolName_ = Trim("" & selRec("TumbolName"))
    aumphurID_ = Trim("" & selRec("AumphurID"))
    aumphurName_ = Trim("" & selRec("AumphurName"))
    provinceID_ = Trim("" & selRec("ProvinceID"))
    provinceName_ = Trim("" & selRec("ProvinceName"))
    zipCode_ = Trim("" & selRec("ZipCode"))
    Country_ = Trim("" & selRec("Country"))
    TelNo_ = Trim("" & selRec("TelNo"))
    faxNo_ = Trim("" & selRec("FaxNo"))
    ContactName_ = Trim("" & selRec("ContactName"))
    isManifestAddr_ = Trim("" & selRec("isManifestAddr"))
    LengthNum_ = Trim("" & selRec("LengthNum"))
    OilNum_ = Trim("" & selRec("OilNum"))
    RateSingle_ = Trim("" & selRec("RateSingle"))
    RateTrailer_ = Trim("" & selRec("RateTrailer"))
    Rate15Q_ = Trim("" & selRec("Rate15Q"))
    Rate24Q_ = Trim("" & selRec("Rate24Q"))
    RatePickup_ = Trim("" & selRec("RatePickup"))
    Rate10_ = Trim("" & selRec("Rate10"))
    RateTrail_ = Trim("" & selRec("RateTrail"))
    RunDay_ = Trim("" & selRec("RunDay"))
    RunHour_ = Trim("" & selRec("RunHour"))
End Sub

