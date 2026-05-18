VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGLABParameterList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim parameterName_ As String
Dim paramCode_ As String
Dim labUnitID_ As String
Dim labUnitName_ As String
Dim colUnit_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get ParameterName() As String
    ParameterName = parameterName_
End Property
Public Property Let ParameterName(ByVal newVal As String)
    If Trim(parameterName_) <> Trim(newVal) Then
        parameterName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ParameterCode() As String
    ParameterCode = paramCode_
End Property
Public Property Let ParameterCode(ByVal newVal As String)
    If Trim(paramCode_) <> Trim(newVal) Then
        paramCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LabUnitID() As String
    LabUnitID = labUnitID_
End Property
Public Property Let LabUnitID(ByVal newVal As String)
    If Trim(labUnitID_) <> Trim(newVal) Then
        labUnitID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LabUnitName() As String
    LabUnitName = labUnitName_
End Property
Public Property Let LabUnitName(ByVal newVal As String)
    If Trim(labUnitName_) <> Trim(newVal) Then
        labUnitName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get UnitLAB() As Collection
Dim tmpS As New BWGSearchManager
    If colUnit_ Is Nothing Then Set colUnit_ = tmpS.UnitLABSearch("LabUnitID='" & id_ & "'")
    Set UnitLAB = colUnit_
End Property
Public Property Set UnitLAB(ByVal newVal As Collection)
    Set colUnit_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryUnitLAB()
    Set colUnit_ = Nothing
End Sub

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
    tmpRec.Open "Select * from tbLABParameterList Where ParameterID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("ParameterID") = id_
    End If
    tmpRec("ParameterName") = Trim(parameterName_)
    tmpRec("ParamCode") = Trim(paramCode_)
    tmpRec("LabUnitID") = Trim(labUnitID_)
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
    DBManager.GetDBConnection.Execute "Delete tbLABParameterList Where ParameterID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("ParameterID")
    parameterName_ = Trim("" & selRec("ParameterName"))
    paramCode_ = Trim("" & selRec("ParamCode"))
    labUnitID_ = Trim("" & selRec("LabUnitID"))
    labUnitName_ = Trim("" & selRec("UnitName"))
End Sub

