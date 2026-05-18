VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGLABTestType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim testTypeName_ As String
Dim testPriority_ As Integer
Dim isChanged_ As Boolean
'
Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get TestTypeName() As String
    TestTypeName = testTypeName_
End Property
Public Property Let TestTypeName(ByVal newVal As String)
    If Trim(testTypeName_) <> Trim(newVal) Then
        testTypeName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TestPriority() As Integer
    TestPriority = testPriority_
End Property
Public Property Let TestPriority(ByVal newVal As Integer)
    If testPriority_ <> newVal Then
        testPriority_ = newVal
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
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbLABTestType Where TestTypeID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("TestTypeID") = id_
    End If
    tmpRec("TestTypeName") = Trim(testTypeName_)
    tmpRec("TestPriority") = testPriority_
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
    DBManager.GetDBConnection.Execute "Delete tbLABTestType Where TestTypeID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("TestTypeID")
    testTypeName_ = Trim("" & selRec("TestTypeName"))
    testPriority_ = selRec("TestPriority")
End Sub

