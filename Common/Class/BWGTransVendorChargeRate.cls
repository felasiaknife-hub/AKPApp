VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGTransVendorChargeRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim carTypeID_ As String
Dim carTypeName_ As String
Dim fromAumphurID_ As String
Dim fromAumphurName_ As String
Dim fromProvinceID_ As String
Dim fromProvinceName_ As String
Dim toAumphurID_ As String
Dim toAumphurName_ As String
Dim toProvinceID_ As String
Dim toProvinceName_ As String
Dim ratePerKM_ As Double
Dim ratePerTrip_ As Double
Dim ratePerVol_ As Double
Dim estDistance_ As Double
Dim lastUpdateDate_ As String
Dim chargeSubID_ As String
Dim custRatePerKM_ As Double
Dim custRatePerTrip_ As Double
Dim custRatePerVol_ As Double

Dim trailerRateKM_ As Double
Dim trailerRateTrip_ As Double
Dim trailerRateVol_ As Double
Dim trailerCustRateKM_ As Double
Dim trailerCustRateTrip_ As Double
Dim trailerCustRateVol_ As Double
Dim isChanged_ As Boolean
Dim ActiveRate_ As Integer
Dim StartDate_ As String
Dim EndDate_ As String
Dim Allowance_ As Double
Dim Dream_ As Double

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get CarTypeID() As String
    CarTypeID = carTypeID_
End Property
Public Property Let CarTypeID(ByVal newval As String)
    If Trim(carTypeID_) <> Trim(newval) Then
        carTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CarTypeName() As String
    CarTypeName = carTypeName_
End Property
Public Property Let CarTypeName(ByVal newval As String)
    If Trim(carTypeName_) <> Trim(newval) Then
        carTypeName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FromAumphurID() As String
    FromAumphurID = fromAumphurID_
End Property
Public Property Let FromAumphurID(ByVal newval As String)
    If Trim(fromAumphurID_) <> Trim(newval) Then
        fromAumphurID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FromAumphurName() As String
    FromAumphurName = fromAumphurName_
End Property
Public Property Let FromAumphurName(ByVal newval As String)
    If Trim(fromAumphurName_) <> Trim(newval) Then
        fromAumphurName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FromProvinceID() As String
    FromProvinceID = fromProvinceID_
End Property
Public Property Let FromProvinceID(ByVal newval As String)
    If Trim(fromProvinceID_) <> Trim(newval) Then
        fromProvinceID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get FromProvinceName() As String
    FromProvinceName = fromProvinceName_
End Property
Public Property Let FromProvinceName(ByVal newval As String)
    If Trim(fromProvinceName_) <> Trim(newval) Then
        fromProvinceName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ToAumphurID() As String
    ToAumphurID = toAumphurID_
End Property
Public Property Let ToAumphurID(ByVal newval As String)
    If Trim(toAumphurID_) <> Trim(newval) Then
        toAumphurID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ToAumphurName() As String
    ToAumphurName = toAumphurName_
End Property
Public Property Let ToAumphurName(ByVal newval As String)
    If Trim(toAumphurName_) <> Trim(newval) Then
        toAumphurName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ToProvinceID() As String
    ToProvinceID = toProvinceID_
End Property
Public Property Let ToProvinceID(ByVal newval As String)
    If Trim(toProvinceID_) <> Trim(newval) Then
        toProvinceID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ToProvinceName() As String
    ToProvinceName = toProvinceName_
End Property
Public Property Let ToProvinceName(ByVal newval As String)
    If Trim(toProvinceName_) <> Trim(newval) Then
        toProvinceName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get RatePerKM() As Double
    RatePerKM = ratePerKM_
End Property
Public Property Let RatePerKM(ByVal newval As Double)
    If ratePerKM_ <> newval Then
        ratePerKM_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get CustRatePerKM() As Double
    CustRatePerKM = custRatePerKM_
End Property
Public Property Let CustRatePerKM(ByVal newval As Double)
    If custRatePerKM_ <> newval Then
        custRatePerKM_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get CustRatePerTrip() As Double
    CustRatePerTrip = custRatePerTrip_
End Property
Public Property Let CustRatePerTrip(ByVal newval As Double)
    If custRatePerTrip_ <> newval Then
        custRatePerTrip_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get RatePerTrip() As Double
    RatePerTrip = ratePerTrip_
End Property
Public Property Let RatePerTrip(ByVal newval As Double)
    If ratePerTrip_ <> newval Then
        ratePerTrip_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get RatePerVol() As Double
    RatePerVol = ratePerVol_
End Property
Public Property Let RatePerVol(ByVal newval As Double)
    If ratePerVol_ <> newval Then
        ratePerVol_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get CustRatePerVol() As Double
    CustRatePerVol = custRatePerVol_
End Property
Public Property Let CustRatePerVol(ByVal newval As Double)
    If custRatePerVol_ <> newval Then
        custRatePerVol_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get EstDistance() As Double
    EstDistance = estDistance_
End Property
Public Property Let EstDistance(ByVal newval As Double)
    If estDistance_ <> newval Then
        estDistance_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get LastUpdateDate() As String
    LastUpdateDate = lastUpdateDate_
End Property
Public Property Let LastUpdateDate(ByVal newval As String)
    If Trim(lastUpdateDate_) <> Trim(newval) Then
        lastUpdateDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get ChargeSubID() As String
    ChargeSubID = chargeSubID_
End Property
Public Property Let ChargeSubID(ByVal newval As String)
    If Trim(chargeSubID_) <> Trim(newval) Then
        chargeSubID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TrailerRateKM() As Double
    TrailerRateKM = trailerRateKM_
End Property
Public Property Let TrailerRateKM(ByVal newval As Double)
    If trailerRateKM_ <> newval Then
        trailerRateKM_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TrailerRateTrip() As Double
    TrailerRateTrip = trailerRateTrip_
End Property
Public Property Let TrailerRateTrip(ByVal newval As Double)
    If trailerRateTrip_ <> newval Then
        trailerRateTrip_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get Allowance() As Double
    Allowance = Allowance_
End Property
Public Property Let Allowance(ByVal newval As Double)
    If Allowance_ <> newval Then
        Allowance_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get Dream() As Double
    Dream = Dream_
End Property
Public Property Let Dream(ByVal newval As Double)
       If Dream_ <> newval Then
        Dream_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TrailerRateVol() As Double
    TrailerRateVol = trailerRateVol_
End Property
Public Property Let TrailerRateVol(ByVal newval As Double)
    If trailerRateVol_ <> newval Then
        trailerRateVol_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TrailerCustRateKM() As Double
    TrailerCustRateKM = trailerCustRateKM_
End Property
Public Property Let TrailerCustRateKM(ByVal newval As Double)
    If trailerCustRateKM_ <> newval Then
        trailerCustRateKM_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TrailerCustRateTrip() As Double
    TrailerCustRateTrip = trailerCustRateTrip_
End Property
Public Property Let TrailerCustRateTrip(ByVal newval As Double)
    If trailerCustRateTrip_ <> newval Then
        trailerCustRateTrip_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TrailerCustRateVol() As Double
    TrailerCustRateVol = trailerCustRateVol_
End Property
Public Property Let TrailerCustRateVol(ByVal newval As Double)
    If trailerCustRateVol_ <> newval Then
        trailerCustRateVol_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get ActiveRate() As Integer
    ActiveRate = ActiveRate_
End Property
Public Property Let ActiveRate(ByVal newval As Integer)
    If ActiveRate_ <> newval Then
        ActiveRate_ = newval
    End If
End Property

Public Property Get StartDate() As String
    StartDate = StartDate_
End Property
Public Property Let StartDate(ByVal newval As String)
    If Trim(StartDate_) <> Trim(newval) Then
        StartDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get EndDate() As String
    EndDate = EndDate_
End Property
Public Property Let EndDate(ByVal newval As String)
    If Trim(EndDate_) <> Trim(newval) Then
        EndDate_ = Trim(newval)
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
    sqlStr = "SELECT * FROM vw_TransportVendorChargeRate WHERE ChargeRateID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
' *** ยกเลิกใช้ การบันทึกค่าขนส่งของ Vendor
'    GetDBConnection
'    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
'    If tmpRec.EOF Then
'        tmpRec.AddNew
'        tmpRec("ChargeRateID") = id_
'    End If
'    tmpRec("CarTypeID") = Trim(carTypeID_)
'    tmpRec("FromAumphurID") = Trim(fromAumphurID_)
'    tmpRec("ToAumphurID") = Trim(toAumphurID_)
'    tmpRec("RatePerKM") = Trim(ratePerKM_)
'    tmpRec("RatePerTrip") = Trim(ratePerTrip_)
'    tmpRec("RatePerVol") = Trim(ratePerVol_)
'    tmpRec("EstDistance") = Trim(estDistance_)
'    tmpRec("LastUpdateDate") = Trim(lastUpdateDate_)
'    tmpRec("ChargeSubID") = Trim(chargeSubID_)
'    tmpRec("CustRatePerKM") = custRatePerKM_
'    tmpRec("CustRatePerVol") = custRatePerVol_
'
'    tmpRec("TrailerRateKM") = trailerRateKM_
'    tmpRec("TrailerRateTrip") = trailerRateTrip_
'    tmpRec("TrailerRateVol") = trailerRateVol_
'    tmpRec("TrailerCustRateKM") = trailerCustRateKM_
'    tmpRec("TrailerCustRateVol") = trailerCustRateVol_
'
'    Select Case ActiveRate_
'        Case 0
'            tmpRec("CustRatePerTrip") = custRatePerTrip_
'            tmpRec("TrailerCustRateTrip") = trailerCustRateTrip_
'            tmpRec("CurCustRatePerTrip") = custRatePerTrip_
'            tmpRec("CurTrailerCustRateTrip") = trailerCustRateTrip_
'        Case 1
'            tmpRec("CustRatePerTrip1") = custRatePerTrip_
'            tmpRec("TrailerCustRateTrip1") = trailerCustRateTrip_
'            tmpRec("CurCustRatePerTrip") = custRatePerTrip_
'            tmpRec("CurTrailerCustRateTrip") = trailerCustRateTrip_
'        Case 2
'            tmpRec("CustRatePerTrip2") = custRatePerTrip_
'            tmpRec("TrailerCustRateTrip2") = trailerCustRateTrip_
'            tmpRec("CurCustRatePerTrip") = custRatePerTrip_
'            tmpRec("CurTrailerCustRateTrip") = trailerCustRateTrip_
'        Case 3
'            tmpRec("CustRatePerTrip3") = custRatePerTrip_
'            tmpRec("TrailerCustRateTrip3") = trailerCustRateTrip_
'            tmpRec("CurCustRatePerTrip") = custRatePerTrip_
'            tmpRec("CurTrailerCustRateTrip") = trailerCustRateTrip_
'    End Select
'    tmpRec.Update
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
'On Error GoTo delErr:
'    DBConnExc "Delete tbTransVendorChargeRate WHERE ChargeRateID='" & id_ & "'"
'    Delete = ""
'    Exit Function
'delErr:
'    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("ChargeRateID"))
    carTypeID_ = Trim("" & selRec("CarTypeID"))
    If isExist(selRec, "FromAumphurID") Then fromAumphurID_ = Trim("" & selRec("FromAumphurID"))
    If isExist(selRec, "ToAumphurID") Then toAumphurID_ = Trim("" & selRec("ToAumphurID"))
    If isExist(selRec, "RatePerTrip") Then ratePerTrip_ = IIf(IsNull(selRec!RatePerTrip), 0, selRec!RatePerTrip)
    If isExist(selRec, "TrailerRateTrip") Then trailerRateTrip_ = IIf(IsNull(selRec!TrailerRateTrip), 0, selRec!TrailerRateTrip)
    If isExist(selRec, "StartDate") Then StartDate_ = Trim("" & selRec("StartDate"))
    If isExist(selRec, "EndDate") Then EndDate_ = Trim("" & selRec("EndDate"))
    If isExist(selRec, "Allowance") Then Allowance_ = IIf(IsNull(selRec!Allowance), 0, selRec!Allowance)
    If isExist(selRec, "Dream") Then Dream_ = IIf(IsNull(selRec!Dream), 0, selRec!Dream)
End Sub
