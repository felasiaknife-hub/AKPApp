VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCancelRemark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim RemID_ As String
Dim Remdesc_ As String

Public Property Get RemID() As String
    RemID = RemID_
End Property
Public Property Let RemID(ByVal newVal As String)
    RemID_ = Trim(newVal)
End Property

Public Property Get Remdesc() As String
    Remdesc = Remdesc_
End Property
Public Property Let Remdesc(ByVal newVal As String)
    If Trim(Remdesc_) <> Trim(newVal) Then
        Remdesc_ = Trim(newVal)
    End If
End Property
'Public Sub Save()
'On Error GoTo ErrD
'Dim tmpRec As ADODB.Recordset
'Dim sqlStr As String
'    Set tmpRec = New ADODB.Recordset
'    sqlStr = "SELECT * FROM tbGeneralUnit WHERE GenUnitID='" & id_ & "'"
'    GetDBConnection
'    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
'    If tmpRec.EOF Then
'        tmpRec.AddNew
'        tmpRec("GenUnitID") = id_
'    End If
'    tmpRec("UnitName") = Trim(unitName_)
'    tmpRec.Update
''    tmpRec.Close
'    Set tmpRec = Nothing
'    isChanged_ = False
'    CloseDBConnection
'    Exit Sub
'ErrD:
'    GetMsg "", vbCritical
'    Exit Sub
'End Sub
'
'Public Function Delete() As String
'On Error GoTo delErr:
'    DBConnExc "Delete tbGeneralUnit WHERE GenUnitID='" & id_ & "'"
'    Delete = ""
'    Exit Function
'delErr:
'    Delete = err.Description
'End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    RemID_ = selRec("RemID")
    Remdesc_ = Trim("" & selRec("Remdesc"))
End Sub

