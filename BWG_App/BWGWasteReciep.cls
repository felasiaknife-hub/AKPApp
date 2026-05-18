VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGWasteReciep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim MenifestID_ As String
Dim wasteDataID_ As String
Dim DocNo_ As String
Dim workDate_ As String
Dim CustomerName_ As String
Dim wasteName_ As String
Dim Weight_ As String
Dim treatmentRate_ As String
Dim treatmentUnitID_ As String
Dim treatmentCharge_ As String
Dim CarRegisID_ As String
Dim transportFee_ As String
Dim isSendakp_ As String
Dim isclosed_ As String
Public Property Get menifestID() As String
    menifestID = MenifestID_
End Property
Public Property Let menifestID(ByVal newVal As String)
    MenifestID_ = Trim(newVal)
End Property
Public Property Get wastedataID() As String
    wastedataID = wasteDataID_
End Property
Public Property Let wastedataID(ByVal newVal As String)
    If Trim(wasteDataID_) <> Trim(newVal) Then
        wasteDataID_ = Trim(newVal)
    End If
End Property
Public Property Get docNo() As String
    docNo = DocNo_
End Property
Public Property Let docNo(ByVal newVal As String)
    If Trim(DocNo_) <> Trim(newVal) Then
        DocNo_ = Trim(newVal)
    End If
End Property
Public Property Get WorkDate() As String
    WorkDate = workDate_
End Property
Public Property Let WorkDate(ByVal newVal As String)
    If Trim(workDate_) <> Trim(newVal) Then
        workDate_ = Trim(newVal)
    End If
End Property
Public Property Get CustomerName() As String
    CustomerName = CustomerName_
End Property
Public Property Let CustomerName(ByVal newVal As String)
    If Trim(CustomerName_) <> Trim(newVal) Then
        CustomerName_ = Trim(newVal)
    End If
End Property
Public Property Get WasteName() As String
    WasteName = wasteName_
End Property
Public Property Let WasteName(ByVal newVal As String)
    If Trim(wasteName_) <> Trim(newVal) Then
        wasteName_ = Trim(newVal)
    End If
End Property
Public Property Get Weight() As Integer
    Weight = Weight_
End Property
Public Property Let Weight(ByVal newVal As Integer)
    If Weight_ <> newVal Then
        Weight_ = newVal
    End If
End Property
Public Property Get TreatmentRate() As Integer
    TreatmentRate = treatmentRate_
End Property
Public Property Let TreatmentRate(ByVal newVal As Integer)
    If treatmentRate_ <> newVal Then
        treatmentRate_ = newVal
    End If
End Property
Public Property Get TreatmentUnitID() As String
    TreatmentUnitID = treatmentUnitID_
End Property
Public Property Let TreatmentUnitID(ByVal newVal As String)
    If treatmentUnitID_ <> newVal Then
        treatmentUnitID_ = newVal
    End If
End Property
Public Property Get TreatmentCharge() As Double
    TreatmentCharge = treatmentCharge_
End Property
Public Property Let TreatmentCharge(ByVal newVal As Double)
    If treatmentCharge_ <> newVal Then
        treatmentCharge_ = newVal
    End If
End Property
Public Property Get CarRegisID() As String
    CarRegisID = CarRegisID_
End Property
Public Property Let CarRegisID(ByVal newVal As String)
    If CarRegisID_ <> newVal Then
        CarRegisID_ = newVal
    End If
End Property
Public Property Get TransportFee() As Double
    TransportFee = transportFee_
End Property
Public Property Let TransportFee(ByVal newVal As Double)
    If transportFee_ <> newVal Then
        transportFee_ = newVal
    End If
End Property
Public Property Get isSendakp() As String
    isSendakp = isSendakp_
End Property
Public Property Let isSendakp(ByVal newVal As String)
    If isSendakp_ <> newVal Then
        isSendakp_ = newVal
    End If
End Property
Public Property Get isclosed() As String
    isclosed = isclosed_
End Property
Public Property Let isclosed(ByVal newVal As String)
    If isclosed_ <> newVal Then
        isclosed_ = newVal
    End If
End Property

Public Sub Save()
'On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpRec2 As ADODB.Recordset
Dim sqlStr2 As String
Dim tmpList As ListItem
Dim i As Integer
Dim TSFee As Integer
Dim RS As ADODB.Recordset
    sqlStr2 = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM vw_Wastereciept3 WHERE MenifestID='" & MenifestID_ & "'"
    Set RS = New ADODB.Recordset
    Set RS = GetRS(sqlStr2, adOpenForwardOnly, adLockReadOnly)
If RS.EOF Then

End If
'    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbReportAKP WHERE MenifestID='" & MenifestID_ & "' and wastedataid='" & WasteDataID_ & "' and docno='" & DocNo_ & "' "
'    Set tmpRec = New ADODB.Recordset
'    GetDBConnection
'    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
'    rs.MoveFirst
'    If tmpRec.EOF Then
'        For i = 1 To rs.RecordCount
'            tmpRec.AddNew
'            tmpRec("MenifestID") = rs("MenifestID")
'            tmpRec("WasteDataID") = rs("WasteDataID")
'            tmpRec("DocNo") = rs("DocNo")
'            tmpRec("WorkDate") = rs("WorkDate")
'            tmpRec("CustomerName") = rs("CustomerName")
'            tmpRec("WasteName") = rs("WasteName")
'            tmpRec("Weight") = rs("Weight")
'            tmpRec("WasteRate") = rs("TreatmentRate")
'            tmpRec("Unit") = rs("TreatmentUnitID")
'            tmpRec("TreatmentPrice") = rs("Treatmentcharge")
'            tmpRec("TruckTypeID") = rs("CarRegisID")
'                If i > 1 Then
'                    TSFee = 0
'                Else
'                    TSFee = rs("TransportFee")
'                End If
'            tmpRec("TranspostFree") = TSFee
'            tmpRec.Update
'                DBConnExc "UPDATE tbJobDataDetail Set isSendAKP = 'Y' Where JobDetID = '" & rs("JobDetID") & "'"
'                DBConnExc "UPDATE tbMenifestWasteAdded Set isSendAKP = 'Y' Where MenifestID = '" & rs("MenifestID") & "' And WasteDataID = '" & rs("WasteDataID") & "'"
'                DBConnExc "UPDATE tbReceiptDetail Set isSendAKP = 'Y' Where ReceiptID = '" & rs("MenifestID") & "' and WasteID = '" & rs("WasteDataID") & "'"
'                rs.MoveNext
'        Next
''    tmpRec.Close
'        Set tmpRec = Nothing
'        CloseDBConnection
'        Exit Sub
'    Else
'        Exit Sub
'    End If
'ErrD:
'    GetMsg "", vbCritical
    Exit Sub
End Sub
