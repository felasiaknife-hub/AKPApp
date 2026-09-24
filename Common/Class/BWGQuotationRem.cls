VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGQuotationRem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim QuotationID_ As String
Dim remDesc_ As String
Dim itemNo_ As Integer
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get QuotationID() As String
    QuotationID = QuotationID_
End Property
Public Property Let QuotationID(ByVal newVal As String)
    If Trim(QuotationID_) <> Trim(newVal) Then
        QuotationID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ItemNo() As Integer
    ItemNo = itemNo_
End Property
Public Property Let ItemNo(ByVal newVal As Integer)
    If itemNo_ <> newVal Then
        itemNo_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get RemDesc() As String
    RemDesc = remDesc_
End Property
Public Property Let RemDesc(ByVal newVal As String)
    If Trim(remDesc_) <> Trim(newVal) Then
        remDesc_ = Trim(newVal)
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
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbQuotationRemark WHERE QRemarkID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("QRemarkID") = id_
    End If
    tmpRec("QuotationID") = Trim(QuotationID_)
    tmpRec("RemDesc") = Trim(remDesc_)
    tmpRec("ItemNo") = itemNo_
    tmpRec("MarkAsDelete") = ""
    
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    
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
    DBConnExc "Delete tbQuotationRemark WHERE QRemarkID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("QRemarkID")
    QuotationID_ = Trim("" & selRec("QuotationID"))
    remDesc_ = Trim("" & selRec("RemDesc"))
    If IsNull(selRec("ItemNo")) = False Then itemNo_ = selRec("ItemNo")
End Sub
