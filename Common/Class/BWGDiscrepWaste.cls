VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGDiscrepWaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim MenifestID_ As String
Dim oldMenifestNo_ As String
Dim oldMenifestDocNo_ As String
Dim discrepDesc_ As String
Dim discrepQty_ As Double
Dim actionTaken_ As String
Dim newWasteDataID_ As String
Dim newWasteName_ As String
Dim newWasteNo_ As String
Dim reasonAction_ As String
Dim returnDate_ As String
Dim returnMenifestID_ As String
Dim returnMenifestNo_ As String
Dim returnMenifestDocNo_ As String
Dim moneyToCharge_ As Double
Dim tsdfID_ As String
Dim tsdfName_ As String
Dim wasteType_ As String
Dim qtyUnitID_ As String
Dim qtyUnitName_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get menifestID() As String
    menifestID = MenifestID_
End Property
Public Property Let menifestID(ByVal newVal As String)
    If Trim(MenifestID_) <> Trim(newVal) Then
        MenifestID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get OldMenifestNo() As String
    OldMenifestNo = oldMenifestNo_
End Property
Public Property Let OldMenifestNo(ByVal newVal As String)
    If Trim(oldMenifestNo_) <> Trim(newVal) Then
        oldMenifestNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get OldMenifestDocNo() As String
    OldMenifestDocNo = oldMenifestDocNo_
End Property
Public Property Let OldMenifestDocNo(ByVal newVal As String)
    If Trim(oldMenifestDocNo_) <> Trim(newVal) Then
        oldMenifestDocNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DiscrepDesc() As String
    DiscrepDesc = discrepDesc_
End Property
Public Property Let DiscrepDesc(ByVal newVal As String)
    If Trim(discrepDesc_) <> Trim(newVal) Then
        discrepDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DiscrepQty() As Double
    DiscrepQty = discrepQty_
End Property
Public Property Let DiscrepQty(ByVal newVal As Double)
    If discrepQty_ <> newVal Then
        discrepQty_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get ActionTaken() As String
    ActionTaken = actionTaken_
End Property
Public Property Let ActionTaken(ByVal newVal As String)
    If Trim(actionTaken_) <> Trim(newVal) Then
        actionTaken_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get NewWasteDataID() As String
    NewWasteDataID = newWasteDataID_
End Property
Public Property Let NewWasteDataID(ByVal newVal As String)
    If Trim(newWasteDataID_) <> Trim(newVal) Then
        newWasteDataID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get NewWasteName() As String
    NewWasteName = newWasteName_
End Property
Public Property Let NewWasteName(ByVal newVal As String)
    If Trim(newWasteName_) <> Trim(newVal) Then
        newWasteName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get NewWasteNo() As String
    NewWasteNo = newWasteNo_
End Property
Public Property Let NewWasteNo(ByVal newVal As String)
    If Trim(newWasteNo_) <> Trim(newVal) Then
        newWasteNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ReasonAction() As String
    ReasonAction = reasonAction_
End Property
Public Property Let ReasonAction(ByVal newVal As String)
    If Trim(reasonAction_) <> Trim(newVal) Then
        reasonAction_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ReturnDate() As String
    ReturnDate = returnDate_
End Property
Public Property Let ReturnDate(ByVal newVal As String)
    If Trim(returnDate_) <> Trim(newVal) Then
        returnDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ReturnMenifestID() As String
    ReturnMenifestID = returnMenifestID_
End Property
Public Property Let ReturnMenifestID(ByVal newVal As String)
    If Trim(returnMenifestID_) <> Trim(newVal) Then
        returnMenifestID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ReturnMenifestNo() As String
    ReturnMenifestNo = returnMenifestNo_
End Property
Public Property Let ReturnMenifestNo(ByVal newVal As String)
    If Trim(returnMenifestNo_) <> Trim(newVal) Then
        returnMenifestNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ReturnMenifestDocNo() As String
    ReturnMenifestDocNo = returnMenifestDocNo_
End Property
Public Property Let ReturnMenifestDocNo(ByVal newVal As String)
    If Trim(returnMenifestDocNo_) <> Trim(newVal) Then
        returnMenifestDocNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MoneyToCharge() As Double
    MoneyToCharge = moneyToCharge_
End Property
Public Property Let MoneyToCharge(ByVal newVal As Double)
    If moneyToCharge_ <> newVal Then
        moneyToCharge_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TSDFID() As String
    TSDFID = tsdfID_
End Property
Public Property Let TSDFID(ByVal newVal As String)
    If Trim(tsdfID_) <> Trim(newVal) Then
        tsdfID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFName() As String
    TSDFName = tsdfName_
End Property
Public Property Let TSDFName(ByVal newVal As String)
    If Trim(tsdfName_) <> Trim(newVal) Then
        tsdfName_ = Trim(newVal)
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

Public Property Get QtyUnitID() As String
    QtyUnitID = qtyUnitID_
End Property
Public Property Let QtyUnitID(ByVal newVal As String)
    If Trim(qtyUnitID_) <> Trim(newVal) Then
        qtyUnitID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get QtyUnitName() As String
    QtyUnitName = qtyUnitName_
End Property
Public Property Let QtyUnitName(ByVal newVal As String)
    If Trim(qtyUnitName_) <> Trim(newVal) Then
        qtyUnitName_ = Trim(newVal)
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
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbMenifestDiscrep WHERE MenifDiscrepID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("MenifDiscrepID") = id_
    End If
    tmpRec("MenifestID") = Trim(MenifestID_)
    tmpRec("DiscrepDesc") = Trim(discrepDesc_)
    tmpRec("DiscrepQty") = discrepQty_
    tmpRec("ActionTaken") = Trim(actionTaken_)
    tmpRec("NewWasteDataID") = Trim(newWasteDataID_)
    tmpRec("ReasonOfAction") = Trim(reasonAction_)
    tmpRec("ReturnDate") = Trim(returnDate_)
    tmpRec("ReturnMenifestID") = Trim(returnMenifestID_)
    tmpRec("MoneyToCharge") = moneyToCharge_
    tmpRec("TSDFID") = Trim(tsdfID_)
    tmpRec("WasteType") = Trim(wasteType_)
    tmpRec("QtyUnitID") = Trim(qtyUnitID_)
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
    DBConnExc "Delete tbMenifestDiscrep WHERE MenifDiscrepID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("MenifDiscrepID"))
    MenifestID_ = Trim("" & selRec("MenifestID"))
    oldMenifestNo_ = Trim("" & selRec("OldMenifestNo"))
    oldMenifestDocNo_ = Trim("" & selRec("OldDocNo"))
    discrepDesc_ = Trim("" & selRec("DiscrepDesc"))
    discrepQty_ = selRec("DiscrepQty")
    actionTaken_ = Trim("" & selRec("ActionTaken"))
    newWasteDataID_ = Trim("" & selRec("NewWasteDataID"))
    newWasteName_ = Trim("" & selRec("NewWasteName"))
    newWasteNo_ = Trim("" & selRec("NewWasteNo"))
    reasonAction_ = Trim("" & selRec("ReasonOfAction"))
    returnDate_ = Trim("" & selRec("ReturnDate"))
    returnMenifestID_ = Trim("" & selRec("ReturnMenifestID"))
    returnMenifestNo_ = Trim("" & selRec("NewMenifestNo"))
    returnMenifestDocNo_ = Trim("" & selRec("NewDocNo"))
    moneyToCharge_ = selRec("MoneyToCharge")
    tsdfID_ = Trim("" & selRec("TSDFID"))
    tsdfName_ = Trim("" & selRec("TSDFName"))
    wasteType_ = Trim("" & selRec("WasteType"))
    qtyUnitID_ = Trim("" & selRec("QtyUnitID"))
    qtyUnitName_ = Trim("" & selRec("UnitName"))
End Sub
