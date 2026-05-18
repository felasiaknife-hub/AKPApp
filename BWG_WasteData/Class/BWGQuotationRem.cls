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
Dim quotationID_ As String
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
    QuotationID = quotationID_
End Property
Public Property Let QuotationID(ByVal newVal As String)
    If Trim(quotationID_) <> Trim(newVal) Then
        quotationID_ = Trim(newVal)
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
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbQuotationRemark Where QRemarkID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("QRemarkID") = id_
    End If
    tmpRec("QuotationID") = Trim(quotationID_)
    tmpRec("RemDesc") = Trim(remDesc_)
    tmpRec("ItemNo") = itemNo_
    tmpRec("MarkAsDelete") = ""
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
    DBManager.GetDBConnection.Execute "Delete tbQuotationRemark Where QRemarkID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("QRemarkID")
    quotationID_ = Trim("" & selRec("QuotationID"))
    remDesc_ = Trim("" & selRec("RemDesc"))
    If IsNull(selRec("ItemNo")) = False Then itemNo_ = selRec("ItemNo")
End Sub
