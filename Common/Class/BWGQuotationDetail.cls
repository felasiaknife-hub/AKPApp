VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGQuotationDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim QuotationID_  As String
Dim wasteDataID_ As String
Dim WasteName_ As String
Dim wasteType_ As String
Dim wasteNo_ As String
Dim wasteEUCode_ As String
Dim yearQty_ As Double
Dim TreatmentRate_ As Double
Dim Remdesc_ As String
Dim loadTumbolID_ As String
Dim loadTumbolName_ As String
Dim loadAumphurID_ As String
Dim loadAumphurName_ As String
Dim loadProvinceID_ As String
Dim loadProvinceName_ As String
Dim billDesc_ As String
Dim TreatmentUnitID_ As String
Dim treatmentUnitName_ As String
Dim wasteQtyUnitID_ As String
Dim wasteQtyUnitName_ As String
Dim rowPos_ As Long
Dim isChanged_ As Boolean
Dim Cost_ As Double
Dim CostUnit_ As String
Dim HeatingValue_ As Double
Dim BurningRate_ As Double
Dim BlendingRate_ As Double
Dim ParameterRate_ As Double
Dim CompanyCode_ As String
Dim WasteStatus_ As String

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get QuotationID() As String
    QuotationID = QuotationID_
End Property
Public Property Let QuotationID(ByVal newval As String)
    If Trim(QuotationID_) <> Trim(newval) Then
        QuotationID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get wastedataID() As String
    wastedataID = wasteDataID_
End Property
Public Property Let wastedataID(ByVal newval As String)
    If Trim(wasteDataID_) <> Trim(newval) Then
        wasteDataID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteName() As String
    WasteName = WasteName_
End Property
Public Property Let WasteName(ByVal newval As String)
    If Trim(WasteName_) <> Trim(newval) Then
        WasteName_ = Trim(newval)
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

Public Property Get WasteNo() As String
    WasteNo = wasteNo_
End Property
Public Property Let WasteNo(ByVal newval As String)
    If Trim(wasteNo_) <> Trim(newval) Then
        wasteNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteEUCode() As String
    WasteEUCode = wasteEUCode_
End Property
Public Property Let WasteEUCode(ByVal newval As String)
    If Trim(wasteEUCode_) <> Trim(newval) Then
        wasteEUCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get YearQty() As Double
    YearQty = yearQty_
End Property
Public Property Let YearQty(ByVal newval As Double)
    If yearQty_ <> newval Then
        yearQty_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get RowPos() As Long
    RowPos = rowPos_
End Property
Public Property Let RowPos(ByVal newval As Long)
    If rowPos_ <> newval Then
        rowPos_ = newval
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

Public Property Get LoadTumbolID() As String
    LoadTumbolID = loadTumbolID_
End Property
Public Property Let LoadTumbolID(ByVal newval As String)
    If Trim(loadTumbolID_) <> Trim(newval) Then
        loadTumbolID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LoadTumbolName() As String
    LoadTumbolName = loadTumbolName_
End Property
Public Property Let LoadTumbolName(ByVal newval As String)
    If Trim(loadTumbolName_) <> Trim(newval) Then
        loadTumbolName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LoadAumphurID() As String
    LoadAumphurID = loadAumphurID_
End Property
Public Property Let LoadAumphurID(ByVal newval As String)
    If Trim(loadAumphurID_) <> Trim(newval) Then
        loadAumphurID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LoadAumphurName() As String
    LoadAumphurName = loadAumphurName_
End Property
Public Property Let LoadAumphurName(ByVal newval As String)
    If Trim(loadAumphurName_) <> Trim(newval) Then
        loadAumphurName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LoadProvinceID() As String
    LoadProvinceID = loadProvinceID_
End Property
Public Property Let LoadProvinceID(ByVal newval As String)
    If Trim(loadProvinceID_) <> Trim(newval) Then
        loadProvinceID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LoadProvinceName() As String
    LoadProvinceName = loadProvinceName_
End Property
Public Property Let LoadProvinceName(ByVal newval As String)
    If Trim(loadProvinceName_) <> Trim(newval) Then
        loadProvinceName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get BillDesc() As String
    BillDesc = billDesc_
End Property
Public Property Let BillDesc(ByVal newval As String)
    If Trim(billDesc_) <> Trim(newval) Then
        billDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentRate() As Double
    TreatmentRate = TreatmentRate_
End Property
Public Property Let TreatmentRate(ByVal newval As Double)
    If TreatmentRate_ <> newval Then
        TreatmentRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentUnitID() As String
    TreatmentUnitID = TreatmentUnitID_
End Property
Public Property Let TreatmentUnitID(ByVal newval As String)
    If Trim(TreatmentUnitID_) <> Trim(newval) Then
        TreatmentUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentUnitName() As String
    TreatmentUnitName = treatmentUnitName_
End Property
Public Property Let TreatmentUnitName(ByVal newval As String)
    If Trim(treatmentUnitName_) <> Trim(newval) Then
        treatmentUnitName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteQtyUnitID() As String
    WasteQtyUnitID = wasteQtyUnitID_
End Property
Public Property Let WasteQtyUnitID(ByVal newval As String)
    If Trim(wasteQtyUnitID_) <> Trim(newval) Then
        wasteQtyUnitID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get WasteQtyUnitName() As String
    WasteQtyUnitName = wasteQtyUnitName_
End Property
Public Property Let WasteQtyUnitName(ByVal newval As String)
    If Trim(wasteQtyUnitName_) <> Trim(newval) Then
        wasteQtyUnitName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get Cost() As Double
    Cost = Cost_
End Property
Public Property Let Cost(ByVal newval As Double)
    If Cost_ <> newval Then
        Cost_ = newval
        isChanged_ = True
    End If
End Property
Public Property Get CostUnit() As String
    CostUnit = CostUnit_
End Property
Public Property Let CostUnit(ByVal newval As String)
    If Trim(CostUnit_) <> Trim(newval) Then
        CostUnit_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get CompanyCode() As String
    CompanyCode = CompanyCode_
End Property
Public Property Let CompanyCode(ByVal newval As String)
    If Trim(CompanyCode_) <> Trim(newval) Then
        CompanyCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get HeatingValue() As Double
Dim rs As New ADODB.Recordset
    Set rs = GetRS("SELECT ISNULL(HeatingValue, 0) AS HeatingValue FROM tbWasteDataCR WHERE WasteDataID = '" & wasteDataID_ & "'", adOpenForwardOnly, adLockReadOnly)
    If Not rs.EOF Then
        HeatingValue_ = rs!HeatingValue
    Else
        HeatingValue_ = 0
    End If
    Set rs = Nothing
    HeatingValue = HeatingValue_
End Property
Public Property Let HeatingValue(ByVal newval As Double)
    If HeatingValue_ <> newval Then
        HeatingValue_ = newval
    End If
End Property

Public Property Get BurningRate() As Double
Dim rs As New ADODB.Recordset
Dim HeatValue As Double
    If TreatmentRate_ <> 0 Then
        BurningRate = TreatmentRate_
    Else
        If HeatingValue_ = 0 Then
            BurningRate = 0
        Else
            If HeatingValue_ >= 1000 And HeatingValue_ <= 2000 Then HeatValue = 1000
            If HeatingValue_ >= 2000 And HeatingValue_ <= 3000 Then HeatValue = 2000
            If HeatingValue_ >= 3000 And HeatingValue_ <= 4000 Then HeatValue = 3000
            If HeatingValue_ >= 4000 And HeatingValue_ <= 5000 Then HeatValue = 4000
            If HeatingValue_ >= 5000 And HeatingValue_ <= 6000 Then HeatValue = 5000
            If HeatingValue_ >= 6000 And HeatingValue_ <= 7000 Then HeatValue = 6000
            If HeatingValue_ >= 7000 And HeatingValue_ <= 8000 Then HeatValue = 7000
            If HeatingValue_ >= 8000 And HeatingValue_ <= 9000 Then HeatValue = 8000
            If HeatingValue_ >= 9000 And HeatingValue_ <= 10000 Then HeatValue = 9000
            If HeatingValue_ > 10000 Then HeatValue = 10000
            Set rs = GetRS("SELECT * FROM tbWasteBurningRate WHERE HeatingValue = " & HeatValue, adOpenForwardOnly, adLockReadOnly)
            If Not rs.EOF Then
                BurningRate_ = rs!BurningRate
            Else
                BurningRate_ = 0
            End If
            Set rs = Nothing
            BurningRate = BurningRate_
        End If
    End If
End Property
Public Property Let BurningRate(ByVal newval As Double)
    If BurningRate_ <> newval Then
        BurningRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get BlendingRate() As Double
Dim rs As New ADODB.Recordset
Dim HeatValue As Double
    If BlendingRate_ <> 0 Then
        BlendingRate = BlendingRate_
    Else
        If HeatingValue_ = 0 Then
            BlendingRate = 0
        Else
            If HeatingValue_ > 1000 And HeatingValue_ <= 2000 Then HeatValue = 1000
            If HeatingValue_ > 2000 And HeatingValue_ <= 3000 Then HeatValue = 2000
            If HeatingValue_ > 3000 And HeatingValue_ <= 4000 Then HeatValue = 3000
            If HeatingValue_ > 4000 And HeatingValue_ <= 5000 Then HeatValue = 4000
            If HeatingValue_ > 5000 And HeatingValue_ <= 6000 Then HeatValue = 5000
            If HeatingValue_ > 6000 And HeatingValue_ <= 7000 Then HeatValue = 6000
            If HeatingValue_ > 7000 And HeatingValue_ <= 8000 Then HeatValue = 7000
            If HeatingValue_ > 8000 And HeatingValue_ <= 9000 Then HeatValue = 8000
            If HeatingValue_ > 9000 And HeatingValue_ <= 10000 Then HeatValue = 9000
            If HeatingValue_ > 10000 Then HeatValue = 10000
            Set rs = GetRS("SELECT * FROM tbWasteBurningRate WHERE HeatingValue = " & HeatValue, adOpenForwardOnly, adLockReadOnly)
            If Not rs.EOF Then
                BlendingRate = rs!BlendingRate
            Else
                BlendingRate = 0
            End If
            Set rs = Nothing
        End If
    End If
    
End Property
Public Property Let BlendingRate(ByVal newval As Double)
    If BlendingRate_ <> newval Then
        BlendingRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get ParameterRate() As Double
Dim rs As New ADODB.Recordset
Dim ParaValue As Double
    If ParameterRate_ <> 0 Then
        ParameterRate = ParameterRate_
    Else
        Set rs = GetRS("SELECT * FROM vw_WasteDataParameterIndex WHERE WasteDataID = '" & wasteDataID_ & "'", adOpenForwardOnly, adLockReadOnly)
        If Not rs.EOF Then
            ParaValue = 0
            Do While Not rs.EOF
                ParaValue = ParaValue + rs!ParameterChargeRate
                rs.MoveNext
            Loop
            ParameterRate_ = ParaValue
        Else
            ParameterRate_ = 0
        End If
        Set rs = Nothing
    End If
    ParameterRate = ParameterRate_
End Property
Public Property Let ParameterRate(ByVal newval As Double)
    If ParameterRate_ <> newval Then
        ParameterRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get WasteStatus() As String
    WasteStatus = Trim(WasteStatus_)
End Property
Public Property Let WasteStatus(ByVal newval As String)
    If Trim(WasteStatus_) <> Trim(newval) Then
        WasteStatus_ = Trim(newval)
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
    sqlStr = "SELECT * FROM tbQuotationDetail WHERE QuotationDetID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("QuotationDetID") = id_
    End If
    tmpRec("QuotationID") = Trim(QuotationID_)
    tmpRec("WasteDataID") = Trim(wasteDataID_)
    tmpRec("YearQty") = Trim(yearQty_)
    tmpRec("TreatmentRate") = Trim(TreatmentRate_)
    tmpRec("RemDesc") = Trim(Remdesc_)
    tmpRec("LoadTumbolID") = Trim(loadTumbolID_)
    tmpRec("BillDesc") = Trim(billDesc_)
    tmpRec("TreatmentUnitID") = Trim(TreatmentUnitID_)
    tmpRec("WasteQtyUnitID") = Trim(wasteQtyUnitID_)
    '********** For AP
    tmpRec!BlendigCharge = BlendingRate_
    tmpRec!MoreFixedCharge = ParameterRate_
    '********** End ***********************
    tmpRec("RowPos") = rowPos_
    tmpRec("MarkAsDelete") = ""
    tmpRec("Cost") = Cost_
    tmpRec("CostUnit") = Trim(CostUnit_)
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
    DBConnExc "Delete tbQuotationDetail WHERE QuotationDetID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("QuotationDetID"))
    QuotationID_ = Trim("" & selRec("QuotationID"))
    wasteDataID_ = Trim("" & selRec("WasteDataID"))
    WasteName_ = Trim("" & selRec("WasteName"))
    wasteNo_ = Trim("" & selRec("WasteNo"))
    wasteEUCode_ = Trim("" & selRec("EUCode"))
    If IsNull(selRec("YearQty")) = False Then yearQty_ = selRec("YearQty")
    If IsNull(selRec("TreatmentRate")) = False Then TreatmentRate_ = selRec("TreatmentRate")
    If IsNull(selRec("RowPos")) = False Then rowPos_ = selRec("RowPos")
    Remdesc_ = Trim("" & selRec("RemDesc"))
    loadTumbolID_ = Trim("" & selRec("LoadTumbolID"))
    loadTumbolName_ = Trim("" & selRec("TumbolName"))
    loadAumphurID_ = Trim("" & selRec("AumphurID"))
    loadAumphurName_ = Trim("" & selRec("AumphurName"))
    loadProvinceID_ = Trim("" & selRec("ProvinceID"))
    loadProvinceName_ = Trim("" & selRec("ProvinceName"))
    TreatmentUnitID_ = Trim("" & selRec("TreatmentUnitID"))
    treatmentUnitName_ = Trim("" & selRec("TreatmentUnitID"))       'Trim("" & selRec("TreatmentUnitName"))
    wasteQtyUnitID_ = Trim("" & selRec("WasteQtyUnitID"))
    wasteQtyUnitName_ = Trim("" & selRec("QtyUnitName"))
    Cost_ = IIf(IsNull(selRec!Cost), 0, selRec!Cost)
    CostUnit_ = Trim("" & selRec("CostUnit"))
    CompanyCode_ = Trim("" & selRec("CompanyCode"))
    
    If isExist(selRec, "WasteType") Then wasteType_ = Trim("" & selRec("WasteType"))
    billDesc_ = Trim("" & selRec("WasteName"))
    
    BlendingRate_ = IIf(IsNull(selRec!BlendigCharge), 0, selRec!BlendigCharge)
    ParameterRate_ = IIf(IsNull(selRec!MoreFixedCharge), 0, selRec!MoreFixedCharge)
    WasteStatus_ = IIf(IsNull(selRec!WasteStatus), "P", selRec!WasteStatus)
End Sub
