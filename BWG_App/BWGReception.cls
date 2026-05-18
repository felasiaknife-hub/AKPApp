VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGReception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim StaffName_ As String
Dim Dept_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get Dept() As String
    Dept = Dept_
End Property
Public Property Let Dept(ByVal newval As String)
    If Trim(Dept_) <> Trim(newval) Then
        Dept_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get StaffName() As String
    StaffName = StaffName_
End Property
Public Property Let StaffName(ByVal newval As String)
    If Trim(StaffName_) <> Trim(newval) Then
        StaffName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    Set tmpRec = New ADODB.Recordset
    sqlStr = "SELECT * FROM tbReception WHERE ID='" & id_ & "'"
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("StaffName") = Trim(StaffName_)
        tmpRec("Dept") = Trim(Dept_)
    End If
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
    DBConnExc "Delete tbReception WHERE ID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("ID")
    StaffName_ = Trim("" & selRec("StaffName"))
    Dept_ = Trim("" & selRec("Dept"))
End Sub

