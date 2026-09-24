VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGEUCode2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim code2_ As String
Dim codeDesc_ As String
Dim euCode1ID_ As String
Dim colEUCode3_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get Code2() As String
    Code2 = code2_
End Property
Public Property Let Code2(ByVal newVal As String)
    If Trim(code2_) <> Trim(newVal) Then
        code2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get EUCode1ID() As String
    EUCode1ID = euCode1ID_
End Property
Public Property Let EUCode1ID(ByVal newVal As String)
    If Trim(euCode1ID_) <> Trim(newVal) Then
        euCode1ID_ = Trim(newVal)
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

Public Property Get EUCode3List() As Collection
Dim tmpS As New BWGSearchManager
    If colEUCode3_ Is Nothing Then Set colEUCode3_ = tmpS.EUCode3Search("EUCode2ID='" & id_ & "'")
    Set EUCode3List = colEUCode3_
End Property
Public Property Set EUCode3List(newVal As Collection)
    Set colEUCode3_ = newVal
    isChanged_ = True
End Property

Public Sub RequeryEUCode3List()
    Set colEUCode3_ = Nothing
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
Dim sqlStr As String
    Set tmpRec = New ADODB.Recordset
    sqlStr = "SELECT * FROM tbEUCode2 WHERE EUCode2ID='" & id_ & "'"
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("EUCode2ID") = id_
    End If
    tmpRec("Code2") = Trim(code2_)
    tmpRec("EUCode1ID") = Trim(euCode1ID_)
    tmpRec("CodeDesc") = Trim(codeDesc_)
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
    DBConnExc "Delete tbEUCode2 WHERE EUCode2ID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("EUCode2ID")
    code2_ = Trim("" & selRec("Code2"))
    euCode1ID_ = Trim("" & selRec("EUCode1ID"))
    codeDesc_ = Trim("" & selRec("CodeDesc"))
End Sub
