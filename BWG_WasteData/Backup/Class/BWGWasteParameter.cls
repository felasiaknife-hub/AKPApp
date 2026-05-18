VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGWasteParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim ParameterID_ As String
Dim ParameterName_  As String
Dim ParameterIndex_ As Double

Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get ParameterID() As String
    ParameterID = ParameterID_
End Property
Public Property Let ParameterID(ByVal newVal As String)
    If Trim(ParameterID_) <> Trim(newVal) Then
        ParameterID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ParameterName() As String
    ParameterName = ParameterName_
End Property
Public Property Let ParameterName(ByVal newVal As String)
    If Trim(ParameterName_) <> Trim(newVal) Then
        ParameterName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ParameterIndex() As String
    ParameterIndex = ParameterIndex_
End Property
Public Property Let ParameterIndex(ByVal newVal As String)
    If Trim(ParameterIndex_) <> Trim(newVal) Then
        ParameterIndex_ = Trim(newVal)
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
Dim x%, tSql As String
    Set tmpRec = New ADODB.Recordset
        tSql = "Select * from tbWasteDataParameter Where WasteDataID='" & id_ & "'" & vbCrLf
        tSql = tSql & "AND ParameterID = '" & ParameterID_ & "'"
        tmpRec.Open tSql, DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec!WasteDataID = id_
        tmpRec!ParameterID = ParameterID_
    End If
    tmpRec!ParameterIndex = ParameterIndex_
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec!WasteDataID
    ParameterID_ = selRec!ParameterID
    ParameterIndex_ = selRec!ParameterIndex
End Sub
