VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGExternalCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim wasteName_ As String
Dim wasteNo_ As String
Dim comRateAmt_ As Double
Dim rateType_ As String
Dim comRatePC_ As Double
Dim rateTypePC_ As String
Dim comRecName_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
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

Public Property Get WasteNo() As String
    WasteNo = wasteNo_
End Property
Public Property Let WasteNo(ByVal newVal As String)
    If Trim(wasteNo_) <> Trim(newVal) Then
        wasteNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ComRateAmt() As Double
    ComRateAmt = comRateAmt_
End Property
Public Property Let ComRateAmt(ByVal newVal As Double)
    If comRateAmt_ <> newVal Then
        comRateAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get ComRatePC() As Double
    ComRatePC = comRatePC_
End Property
Public Property Let ComRatePC(ByVal newVal As Double)
    If comRatePC_ <> newVal Then
        comRatePC_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get RateType() As String
    RateType = rateType_
End Property
Public Property Let RateType(ByVal newVal As String)
    If Trim(rateType_) <> Trim(newVal) Then
        rateType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get RateTypePC() As String
    RateTypePC = rateTypePC_
End Property
Public Property Let RateTypePC(ByVal newVal As String)
    If Trim(rateTypePC_) <> Trim(newVal) Then
        rateTypePC_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ComRecName() As String
    ComRecName = comRecName_
End Property
Public Property Let ComRecName(ByVal newVal As String)
    If Trim(comRecName_) <> Trim(newVal) Then
        comRecName_ = Trim(newVal)
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
    sqlStr = "SELECT * FROM tbExternalCom WHERE WasteDataID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("WasteDataID") = id_
    End If
    tmpRec("ComRateAmt") = Trim(comRateAmt_)
    tmpRec("RateType") = Trim(rateType_)
    tmpRec("ComRatePC") = comRatePC_
    tmpRec("RateTypePC") = Trim(rateTypePC_)
    tmpRec("ComRecName") = Trim(comRecName_)
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
    DBConnExc "Delete tbExternalCom WHERE WasteDataID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
   id_ = selRec("WasteDataID")
    wasteName_ = Trim("" & selRec("WasteName"))
    wasteNo_ = Trim("" & selRec("WasteNo"))
    comRateAmt_ = Trim("" & selRec("ComRateAmt"))
    comRatePC_ = Trim("" & selRec("ComRatePC"))
    rateType_ = Trim("" & selRec("RateType"))
    rateTypePC_ = Trim("" & selRec("RateTypePC"))
    comRecName_ = Trim("" & selRec("ComRecName"))
End Sub

