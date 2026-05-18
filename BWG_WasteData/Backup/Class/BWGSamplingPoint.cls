VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGSamplingPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim pointCode_ As String
Dim pointDesc_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get PointCode() As String
    PointCode = pointCode_
End Property
Public Property Let PointCode(ByVal newVal As String)
    If Trim(pointCode_) <> Trim(newVal) Then
        pointCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get PointDesc() As String
    PointDesc = pointDesc_
End Property
Public Property Let PointDesc(ByVal newVal As String)
    If Trim(pointDesc_) <> Trim(newVal) Then
        pointDesc_ = Trim(newVal)
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
    tmpRec.Open "Select * from tbSamplingPoint Where SamplingPointID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("SamplingPointID") = id_
    End If
    tmpRec("PointCode") = Trim(pointCode_)
    tmpRec("PointDesc") = Trim(pointDesc_)
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
    DBManager.GetDBConnection.Execute "Delete tbSamplingPoint Where SamplingPointID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("SamplingPointID")
    pointCode_ = Trim("" & selRec("PointCode"))
    pointDesc_ = Trim("" & selRec("PointDesc"))
End Sub
