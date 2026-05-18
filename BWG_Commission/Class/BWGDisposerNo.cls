VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGDisposerNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim companyID_ As String
Dim disposerNo_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get CompanyID() As String
    CompanyID = companyID_
End Property
Public Property Let CompanyID(ByVal newVal As String)
    If Trim(companyID_) <> Trim(newVal) Then
        companyID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DisposerNo() As String
    DisposerNo = disposerNo_
End Property
Public Property Let DisposerNo(ByVal newVal As String)
    If Trim(disposerNo_) <> Trim(newVal) Then
        disposerNo_ = Trim(newVal)
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
Dim tmpRec As ADODB.Recordset
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbDisposerNo Where DisposerID='" & id_ & "'", DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("DisposerID") = id_
    End If
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("DisposerNo") = Trim(disposerNo_)
    tmpRec.Update
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbDisposerNo Where DisposerID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("DisposerID")
    companyID_ = Trim("" & selRec("CompanyID"))
    disposerNo_ = Trim("" & selRec("DisposerNo"))
End Sub
