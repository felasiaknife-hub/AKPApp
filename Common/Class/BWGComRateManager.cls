VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGComRateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim hzStartPC_ As Integer
Dim hzEndPC_ As Integer
Dim bonusAmt_  As Double
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get HzStartPC() As Integer
    HzStartPC = hzStartPC_
End Property
Public Property Let HzStartPC(ByVal newVal As Integer)
    If hzStartPC_ <> newVal Then
        hzStartPC_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get HzEndPC() As Integer
    HzEndPC = hzEndPC_
End Property
Public Property Let HzEndPC(ByVal newVal As Integer)
    If hzEndPC_ <> newVal Then
        hzEndPC_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BonusAmt() As Integer
    BonusAmt = bonusAmt_
End Property
Public Property Let BonusAmt(ByVal newVal As Integer)
    If bonusAmt_ <> newVal Then
        bonusAmt_ = newVal
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
    sqlStr = "SELECT * FROM ComRateManager WHERE RateID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("RateID") = id_
    End If
    tmpRec("HZStartPC") = hzStartPC_
    tmpRec("HZEndPC") = hzEndPC_
    tmpRec("BonusAmt") = bonusAmt_
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
    DBConnExc "Delete ComRateManager WHERE RateID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("RateID")
    If IsNull(selRec("HzStartPC")) = False Then hzStartPC_ = Trim("" & selRec("HzStartPC"))
    If IsNull(selRec("HzEndPC")) = False Then hzEndPC_ = Trim("" & selRec("HzEndPC"))
    If IsNull(selRec("BonusAmt")) = False Then bonusAmt_ = Trim("" & selRec("BonusAmt"))
End Sub
