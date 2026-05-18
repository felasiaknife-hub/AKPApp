VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGEUCode1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim code1_ As String
Dim codeDesc_ As String
Dim colEUCode2_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get Code1() As String
    Code1 = code1_
End Property
Public Property Let Code1(ByVal newVal As String)
    If Trim(code1_) <> Trim(newVal) Then
        code1_ = Trim(newVal)
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

Public Property Get EUCode2List() As Collection
Dim tmpS As New BWGSearchManager
    If colEUCode2_ Is Nothing Then Set colEUCode2_ = tmpS.EUCode2Search("EUCode1ID='" & id_ & "'")
    Set EUCode2List = colEUCode2_
End Property
Public Property Set EUCode2List(newVal As Collection)
    Set colEUCode2_ = newVal
    isChanged_ = True
End Property

Public Sub RequeryEUCode2List()
    Set colEUCode2_ = Nothing
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
    tmpRec.Open "Select * from tbEUCode1 Where EUCode1ID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("EUCode1ID") = id_
    End If
    tmpRec("Code1") = Trim(code1_)
    tmpRec("CodeDesc") = Trim(codeDesc_)
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
    DBManager.GetDBConnection.Execute "Delete tbEUCode1 Where EUCode1ID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("EUCode1ID")
    code1_ = Trim("" & selRec("Code1"))
    codeDesc_ = Trim("" & selRec("CodeDesc"))
End Sub
