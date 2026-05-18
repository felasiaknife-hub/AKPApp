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
Dim quotationID_  As String
Dim wasteDataID_ As String
Dim wasteName_ As String
Dim wasteType_ As String
Dim wasteNo_ As String
Dim wasteEUCode_ As String
Dim yearQty_ As Double
Dim treatmentRate_ As Double
Dim remDesc_ As String
Dim loadTumbolID_ As String
Dim loadTumbolName_ As String
Dim loadAumphurID_ As String
Dim loadAumphurName_ As String
Dim loadProvinceID_ As String
Dim loadProvinceName_ As String
Dim billDesc_ As String
Dim treatmentUnitID_ As String
Dim treatmentUnitName_ As String
Dim wasteQtyUnitID_ As String
Dim wasteQtyUnitName_ As String
Dim rowPos_ As Long
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get QuotationID() As String
    QuotationID = quotationID_
End Property
Public Property Let QuotationID(ByVal newVal As String)
    If Trim(quotationID_) <> Trim(newVal) Then
        quotationID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteDataID() As String
    WasteDataID = wasteDataID_
End Property
Public Property Let WasteDataID(ByVal newVal As String)
    If Trim(wasteDataID_) <> Trim(newVal) Then
        wasteDataID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteName() As String
    WasteName = wasteName_
End Property
Public Property Let WasteName(ByVal newVal As String)
    If Trim(wasteName_) <> Trim(newVal) Then
        wasteName_ = Trim(newVal)
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

Public Property Get WasteNo() As String
    WasteNo = wasteNo_
End Property
Public Property Let WasteNo(ByVal newVal As String)
    If Trim(wasteNo_) <> Trim(newVal) Then
        wasteNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteEUCode() As String
    WasteEUCode = wasteEUCode_
End Property
Public Property Let WasteEUCode(ByVal newVal As String)
    If Trim(wasteEUCode_) <> Trim(newVal) Then
        wasteEUCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get YearQty() As Double
    YearQty = yearQty_
End Property
Public Property Let YearQty(ByVal newVal As Double)
    If yearQty_ <> newVal Then
        yearQty_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentRate() As Double
    TreatmentRate = treatmentRate_
End Property
Public Property Let TreatmentRate(ByVal newVal As Double)
    If treatmentRate_ <> newVal Then
        treatmentRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get RowPos() As Long
    RowPos = rowPos_
End Property
Public Property Let RowPos(ByVal newVal As Long)
    If rowPos_ <> newVal Then
        rowPos_ = newVal
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

Public Property Get LoadTumbolID() As String
    LoadTumbolID = loadTumbolID_
End Property
Public Property Let LoadTumbolID(ByVal newVal As String)
    If Trim(loadTumbolID_) <> Trim(newVal) Then
        loadTumbolID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LoadTumbolName() As String
    LoadTumbolName = loadTumbolName_
End Property
Public Property Let LoadTumbolName(ByVal newVal As String)
    If Trim(loadTumbolName_) <> Trim(newVal) Then
        loadTumbolName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LoadAumphurID() As String
    LoadAumphurID = loadAumphurID_
End Property
Public Property Let LoadAumphurID(ByVal newVal As String)
    If Trim(loadAumphurID_) <> Trim(newVal) Then
        loadAumphurID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LoadAumphurName() As String
    LoadAumphurName = loadAumphurName_
End Property
Public Property Let LoadAumphurName(ByVal newVal As String)
    If Trim(loadAumphurName_) <> Trim(newVal) Then
        loadAumphurName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LoadProvinceID() As String
    LoadProvinceID = loadProvinceID_
End Property
Public Property Let LoadProvinceID(ByVal newVal As String)
    If Trim(loadProvinceID_) <> Trim(newVal) Then
        loadProvinceID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LoadProvinceName() As String
    LoadProvinceName = loadProvinceName_
End Property
Public Property Let LoadProvinceName(ByVal newVal As String)
    If Trim(loadProvinceName_) <> Trim(newVal) Then
        loadProvinceName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get BillDesc() As String
    BillDesc = billDesc_
End Property
Public Property Let BillDesc(ByVal newVal As String)
    If Trim(billDesc_) <> Trim(newVal) Then
        billDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentUnitID() As String
    TreatmentUnitID = treatmentUnitID_
End Property
Public Property Let TreatmentUnitID(ByVal newVal As String)
    If Trim(treatmentUnitID_) <> Trim(newVal) Then
        treatmentUnitID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentUnitName() As String
    TreatmentUnitName = treatmentUnitName_
End Property
Public Property Let TreatmentUnitName(ByVal newVal As String)
    If Trim(treatmentUnitName_) <> Trim(newVal) Then
        treatmentUnitName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteQtyUnitID() As String
    WasteQtyUnitID = wasteQtyUnitID_
End Property
Public Property Let WasteQtyUnitID(ByVal newVal As String)
    If Trim(wasteQtyUnitID_) <> Trim(newVal) Then
        wasteQtyUnitID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteQtyUnitName() As String
    WasteQtyUnitName = wasteQtyUnitName_
End Property
Public Property Let WasteQtyUnitName(ByVal newVal As String)
    If Trim(wasteQtyUnitName_) <> Trim(newVal) Then
        wasteQtyUnitName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbQuotationDetail Where QuotationDetID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("QuotationDetID") = id_
    End If
    tmpRec("QuotationID") = Trim(quotationID_)
    tmpRec("WasteDataID") = Trim(wasteDataID_)
    tmpRec("YearQty") = Trim(yearQty_)
    tmpRec("TreatmentRate") = Trim(treatmentRate_)
    tmpRec("RemDesc") = Trim(remDesc_)
    tmpRec("LoadTumbolID") = Trim(loadTumbolID_)
    tmpRec("BillDesc") = Trim(billDesc_)
    tmpRec("TreatmentUnitID") = Trim(treatmentUnitID_)
    tmpRec("WasteQtyUnitID") = Trim(wasteQtyUnitID_)
    tmpRec("RowPos") = rowPos_
    tmpRec("MarkAsDelete") = ""
    tmpRec.Update
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
    DBManager.GetDBConnection.Execute "Delete tbQuotationDetail Where QuotationDetID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("QuotationDetID"))
    quotationID_ = Trim("" & selRec("QuotationID"))
    wasteDataID_ = Trim("" & selRec("WasteDataID"))
    wasteName_ = Trim("" & selRec("WasteName"))
    wasteNo_ = Trim("" & selRec("WasteNo"))
    wasteEUCode_ = Trim("" & selRec("EUCode"))
    If IsNull(selRec("YearQty")) = False Then yearQty_ = selRec("YearQty")
    If IsNull(selRec("TreatmentRate")) = False Then treatmentRate_ = selRec("TreatmentRate")
    If IsNull(selRec("RowPos")) = False Then rowPos_ = selRec("RowPos")
    remDesc_ = Trim("" & selRec("RemDesc"))
    loadTumbolID_ = Trim("" & selRec("LoadTumbolID"))
    loadTumbolName_ = Trim("" & selRec("TumbolName"))
    loadAumphurID_ = Trim("" & selRec("AumphurID"))
    loadAumphurName_ = Trim("" & selRec("AumphurName"))
    loadProvinceID_ = Trim("" & selRec("ProvinceID"))
    loadProvinceName_ = Trim("" & selRec("ProvinceName"))
    billDesc_ = Trim("" & selRec("BillDesc"))
    treatmentUnitID_ = Trim("" & selRec("TreatmentUnitID"))
    treatmentUnitName_ = Trim("" & selRec("TreatmentUnitID"))       'Trim("" & selRec("TreatmentUnitName"))
    wasteQtyUnitID_ = Trim("" & selRec("WasteQtyUnitID"))
    wasteQtyUnitName_ = Trim("" & selRec("QtyUnitName"))
    If isExist(selRec, "WasteType") Then wasteType_ = Trim("" & selRec("WasteType"))
End Sub
