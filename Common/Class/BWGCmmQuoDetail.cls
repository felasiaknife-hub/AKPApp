VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCmmQuoDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim id_ As String
Dim Cmid_ As String
Dim QuotationID_ As String
Dim QuotationNo_ As String
Dim isDates_ As String
Dim QuoStatus_ As String

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get CmID() As String
    CmID = Cmid_
End Property
Public Property Let CmID(ByVal newVal As String)
    Cmid_ = Trim(newVal)
End Property

Public Property Get QuotationID() As String
    QuotationID = QuotationID_
End Property
Public Property Let QuotationID(ByVal newVal As String)
    QuotationID_ = Trim(newVal)
End Property

Public Property Get QuotationNo() As String
    QuotationNo = QuotationNo_
End Property
Public Property Let QuotationNo(ByVal newVal As String)
    QuotationNo_ = Trim(newVal)
End Property

Public Property Get isDateS() As String
    isDateS = isDates_
End Property
Public Property Let isDateS(ByVal newVal As String)
    isDates_ = Trim(newVal)
End Property

Public Property Get QuoStatus() As String
    QuoStatus = QuoStatus_
End Property
Public Property Let QuoStatus(ByVal newVal As String)
    QuoStatus_ = Trim(newVal)
End Property

Public Sub Save()
    Dim tmpRec As ADODB.Recordset
    Dim x%
    Dim sqlStr As String
        GetDBConnection
        sqlStr = "SELECT * FROM tbCommentQuoDetail WHERE CmmDtID ='" & id_ & "' "
        Set tmpRec = New ADODB.Recordset
        tmpRec.open sqlStr, DBConn, adOpenKeyset, adLockOptimistic
        If tmpRec.EOF Then
            tmpRec.AddNew
            tmpRec("CmmDtID") = id_
            tmpRec("QuotationID") = Trim(QuotationID_)
            tmpRec("CmID") = Trim(Cmid_)
        End If
        tmpRec("QuotationNo") = Trim(QuotationNo_)
        tmpRec("IsDates") = Trim(isDates_)
        tmpRec("QuoStatus") = Trim(QuoStatus_)
        tmpRec.Update
End Sub

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("CmmDtID"))
    Cmid_ = Trim("" & selRec("CmID"))
    QuotationID_ = Trim("" & selRec("QuotationID"))
    QuotationNo_ = Trim("" & selRec("QuotationNo"))
    isDates_ = Trim("" & selRec("IsDates"))
    QuoStatus_ = Trim("" & selRec("QuoStatus"))
End Sub
