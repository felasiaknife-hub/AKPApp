VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCmmQuoHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim id_ As String
Dim CmQuoNo_ As String
Dim CustomerID_ As String
Dim companyID_ As String
Dim companyNameTh_ As String
Dim CreateDate_ As String
Dim StartDate_ As String
Dim EndDate_ As String
Dim Comment_ As String
Dim colQuotationCmmDet_  As Collection

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get CmQuoNo() As String
    CmQuoNo = CmQuoNo_
End Property
Public Property Let CmQuoNo(ByVal newVal As String)
    CmQuoNo_ = Trim(newVal)
End Property

Public Property Get CustomerID() As String
    CustomerID = CustomerID_
End Property
Public Property Let CustomerID(ByVal newVal As String)
    CustomerID_ = Trim(newVal)
End Property

Public Property Get CompanyID() As String
    CompanyID = companyID_
End Property
Public Property Let CompanyID(ByVal newVal As String)
    companyID_ = Trim(newVal)
End Property
Public Property Get CompanyNameTH() As String
    CompanyNameTH = companyNameTh_
End Property
Public Property Let CompanyNameTH(ByVal newVal As String)
    companyNameTh_ = Trim(newVal)
End Property

Public Property Get CreateDate() As String
    CreateDate = CreateDate_
End Property
Public Property Let CreateDate(ByVal newVal As String)
    CreateDate_ = Trim(newVal)
End Property

Public Property Get StartDate() As String
    StartDate = StartDate_
End Property
Public Property Let StartDate(ByVal newVal As String)
    StartDate_ = Trim(newVal)
End Property

Public Property Get EndDate() As String
    EndDate = EndDate_
End Property
Public Property Let EndDate(ByVal newVal As String)
    EndDate_ = Trim(newVal)
End Property

Public Property Get Comment() As String
    Comment = Comment_
End Property
Public Property Let Comment(ByVal newVal As String)
    Comment_ = Trim(newVal)
End Property

Public Property Get QuotationCmmDetails() As Collection
Dim tmpS As New BWGSearchManager
Dim i As Integer
    Set colQuotationCmmDet_ = tmpS.QuotationDetailCmmSearch("CmID ='" & id_ & "'")
    Set QuotationCmmDetails = colQuotationCmmDet_
End Property

Public Sub Save()
    Dim tmpRec As ADODB.Recordset
    Dim x%
    Dim sqlStr As String
        sqlStr = "SELECT * FROM tbCommentQuoHeader WHERE CmID ='" & id_ & "'"
        Set tmpRec = New ADODB.Recordset
        GetDBConnection
        tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
        If tmpRec.EOF Then
            id_ = GetGUID
            tmpRec.AddNew
            tmpRec("CmID") = id_
            tmpRec("CreateDate") = TodayDate
        End If
        tmpRec("CmQuoNo") = Trim(CmQuoNo_)
        tmpRec("CustomerID") = Trim(CustomerID_)
        tmpRec("CompanyID") = Trim(companyID_)
        tmpRec("StartDate") = Trim(StartDate_)
        tmpRec("EndDate") = Trim(EndDate_)
        tmpRec("Comment") = Trim(Comment_)
        tmpRec.Update
End Sub

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("CmID"))
    CmQuoNo_ = Trim("" & selRec("CmQuoNo"))
    CustomerID_ = Trim("" & selRec("CustomerID"))
    companyID_ = Trim("" & selRec("CompanyID"))
    companyNameTh_ = Trim("" & selRec("CompanyNameTH"))
    CreateDate_ = Trim("" & selRec("CreateDate"))
    StartDate_ = selRec("StartDate")
    EndDate_ = Trim("" & selRec("EndDate"))
    Comment_ = Trim("" & selRec("Comment"))
End Sub



