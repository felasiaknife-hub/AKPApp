VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMIAumphur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim aumName_ As String
Dim provID_ As String
Dim provinceName_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = newVal
End Property

Public Property Get Name() As String
    Name = aumName_
End Property
Public Property Let Name(ByVal newVal As String)
    If Trim(aumName_) <> Trim(newVal) Then
        aumName_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceID() As String
    ProvinceID = provID_
End Property
Public Property Let ProvinceID(ByVal newVal As String)
    If Trim(provID_) <> Trim(newVal) Then
        provID_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get ProvinceName() As String
    ProvinceName = provinceName_
End Property
Public Property Let ProvinceName(ByVal newVal As String)
    If Trim(provinceName_) <> Trim(newVal) Then
        provinceName_ = newVal
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
Dim tmpRec As New ADODB.Recordset
    tmpRec.Open "Select * from tbAumphur Where AumphurID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("AumphurID") = id_
    End If
    tmpRec("AumphurName") = Trim(aumName_)
    tmpRec("ProvinceID") = Trim(provID_)
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
    DBManager.GetDBConnection.Execute "Delete tbAumphur Where AumphurID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("AumphurID"))
    aumName_ = Trim("" & selRec("AumphurName"))
    provID_ = Trim("" & selRec("ProvinceID"))
    provinceName_ = Trim("" & selRec("ProvinceName"))
End Sub

