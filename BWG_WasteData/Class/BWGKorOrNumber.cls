VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGKorOrNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim wasteDataID_ As String
Dim korOrNumber_ As String
Dim KorDate_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get wastedataID() As String
    wastedataID = wasteDataID_
End Property
Public Property Let wastedataID(ByVal newval As String)
    If Trim(wasteDataID_) <> Trim(newval) Then
        wasteDataID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get KorOrNumber() As String
    KorOrNumber = korOrNumber_
End Property
Public Property Let KorOrNumber(ByVal newval As String)
    If Trim(korOrNumber_) <> Trim(newval) Then
        korOrNumber_ = Trim(newval)
        isChanged_ = True
    End If
End Property
Public Property Get KorDate() As String
    KorDate = KorDate_
End Property
Public Property Let KorDate(ByVal newval As String)
    If Trim(KorDate_) <> Trim(newval) Then
        KorDate_ = Trim(newval)
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
    tmpRec.Open "Select * from tbKorOrNumber Where KorOrID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("KorOrID") = id_
    End If
    tmpRec("WasteDataID") = Trim(wasteDataID_)
    tmpRec("KorOrNumber") = Trim(korOrNumber_)
    tmpRec("KorDate") = Trim(KorDate_)
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
    DBManager.GetDBConnection.Execute "Delete tbKorOrNumber Where KorOrID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("KorOrID")
    wasteDataID_ = Trim("" & selRec("WasteDataID"))
    korOrNumber_ = Trim("" & selRec("KorOrNumber"))
    KorDate_ = Trim("" & selRec("KorDate"))
End Sub
