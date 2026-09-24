VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGTreatmentData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim treatmentDesc_ As String
Dim treatmentRate_ As Double
Dim treatmentUnit_ As String
Dim treatmentRateWeight_ As Double
Dim treatmentUnitWeight_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get treatmentDesc() As String
    treatmentDesc = treatmentDesc_
End Property
Public Property Let treatmentDesc(ByVal newVal As String)
    If Trim(treatmentDesc_) <> Trim(newVal) Then
        treatmentDesc_ = Trim(newVal)
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

Public Property Get TreatmentUnit() As String
    TreatmentUnit = treatmentUnit_
End Property
Public Property Let TreatmentUnit(ByVal newVal As String)
    If Trim(treatmentUnit_) <> Trim(newVal) Then
        treatmentUnit_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentRateWeight() As Double
    TreatmentRateWeight = treatmentRateWeight_
End Property
Public Property Let TreatmentRateWeight(ByVal newVal As Double)
    If treatmentRateWeight_ <> newVal Then
        treatmentRateWeight_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get TreatmentUnitWeight() As String
    TreatmentUnitWeight = treatmentUnitWeight_
End Property
Public Property Let TreatmentUnitWeight(ByVal newVal As String)
    If Trim(treatmentUnitWeight_) <> Trim(newVal) Then
        treatmentUnitWeight_ = Trim(newVal)
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
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbTreatmentData WHERE TreatmentID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("TreatmentID") = id_
    End If
    tmpRec("TreatmentDesc") = Trim(treatmentDesc_)
    tmpRec("TreatmentRate") = treatmentRate_
    tmpRec("TreatmentUnit") = Trim(treatmentUnit_)
    tmpRec("TreatmentRateWeight") = treatmentRateWeight_
    tmpRec("TreatmentUnitWeight") = Trim(treatmentUnitWeight_)
    
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    
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
    DBConnExc "Delete tbTreatmentData WHERE TreatmentID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("TreatmentID")
    treatmentDesc_ = Trim("" & selRec("TreatmentDesc"))
    If IsNull(selRec("TreatmentRate")) = False Then treatmentRate_ = Trim("" & selRec("TreatmentRate"))
    treatmentUnit_ = Trim("" & selRec("TreatmentUnit"))
    If IsNull(selRec("TreatmentRateWeight")) = False Then treatmentRateWeight_ = Trim("" & selRec("TreatmentRateWeight"))
    treatmentUnitWeight_ = Trim("" & selRec("TreatmentUnitWeight"))
End Sub
