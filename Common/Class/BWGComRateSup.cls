VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGComRateSup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim startPC_ As Integer
Dim endPC_ As Integer
Dim comRate_  As Double
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get StartPC() As Integer
    StartPC = startPC_
End Property
Public Property Let StartPC(ByVal newVal As Integer)
    If startPC_ <> newVal Then
        startPC_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get EndPC() As Integer
    EndPC = endPC_
End Property
Public Property Let EndPC(ByVal newVal As Integer)
    If endPC_ <> newVal Then
        endPC_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get ComRate() As Integer
    ComRate = comRate_
End Property
Public Property Let ComRate(ByVal newVal As Integer)
    If comRate_ <> newVal Then
        comRate_ = newVal
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
    sqlStr = "SELECT * FROM ComRateSup WHERE RateID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("RateID") = id_
    End If
    tmpRec("StartPC") = startPC_
    tmpRec("EndPC") = endPC_
    tmpRec("ComRate") = comRate_
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
    DBConnExc "Delete ComRateSup WHERE RateID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("RateID")
    If IsNull(selRec("StartPC")) = False Then startPC_ = Trim("" & selRec("StartPC"))
    If IsNull(selRec("EndPC")) = False Then endPC_ = Trim("" & selRec("EndPC"))
    If IsNull(selRec("ComRate")) = False Then comRate_ = Trim("" & selRec("ComRate"))
End Sub
