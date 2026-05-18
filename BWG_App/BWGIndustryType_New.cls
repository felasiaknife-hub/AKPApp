VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGIndustryType_New"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim TypeID_ As String
Dim TypeDesc_ As String
Dim isChanged_ As Boolean
Dim TypeCode_ As String

Public Property Get TypeID() As String
    TypeID = TypeID_
End Property
Public Property Let TypeID(ByVal newval As String)
    TypeID_ = Trim(newval)
End Property

Public Property Get TypeDesc() As String
    TypeDesc = TypeDesc_
End Property
Public Property Let TypeDesc(ByVal newval As String)
    If Trim(TypeDesc_) <> Trim(newval) Then
        TypeDesc_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get TypeCode() As String
    TypeCode = TypeCode_
End Property
Public Property Let TypeCode(ByVal newval As String)
    If Trim(TypeCode_) <> Trim(newval) Then
        TypeCode_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
    Set tmpRec = New ADODB.Recordset
Dim sqlStr As String
    Set tmpRec = New ADODB.Recordset
    sqlStr = "SELECT * FROM tbIndustryType WHERE IndustryTypeID='" & TypeID_ & "'"
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("TypeCode") = TypeID_
    End If
    tmpRec("TypeDesc") = Trim(TypeDesc_)
    tmpRec.Update
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub


Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    TypeID_ = selRec("TypeID")
    TypeDesc_ = Trim("" & selRec("TypeDesc"))
    TypeCode_ = selRec("TypeCode")
End Sub



