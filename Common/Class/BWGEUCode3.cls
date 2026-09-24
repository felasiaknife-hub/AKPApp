VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGEUCode3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim code3_ As String
Dim codeDesc_ As String
Dim haType_ As String
Dim euCode2ID_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get Code3() As String
    Code3 = code3_
End Property
Public Property Let Code3(ByVal newVal As String)
    If Trim(code3_) <> Trim(newVal) Then
        code3_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get EUCode2ID() As String
    EUCode2ID = euCode2ID_
End Property
Public Property Let EUCode2ID(ByVal newVal As String)
    If Trim(euCode2ID_) <> Trim(newVal) Then
        euCode2ID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CodeDesc() As String
    CodeDesc = codeDesc_
End Property
Public Property Let CodeDesc(ByVal newVal As String)
    If Trim(codeDesc_) <> Trim(newVal) Then
        codeDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get HAType() As String
    HAType = haType_
End Property
Public Property Let HAType(ByVal newVal As String)
    If Trim(haType_) <> Trim(newVal) Then
        haType_ = Trim(newVal)
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
    Set tmpRec = New ADODB.Recordset
    sqlStr = "SELECT * FROM tbEUCode3 WHERE EUCode3ID='" & id_ & "'"
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("EUCode3ID") = id_
    End If
    tmpRec("Code3") = Trim(code3_)
    tmpRec("EUCode2ID") = Trim(euCode2ID_)
    tmpRec("CodeDesc") = Trim(codeDesc_)
    tmpRec("HAType") = Trim(haType_)
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
    DBConnExc "Delete tbEUCode3 WHERE EUCode3ID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("EUCode3ID")
    code3_ = Trim("" & selRec("Code3"))
    euCode2ID_ = Trim("" & selRec("EUCode2ID"))
    codeDesc_ = Trim("" & selRec("CodeDesc"))
    haType_ = Trim("" & selRec("HAType"))
End Sub
