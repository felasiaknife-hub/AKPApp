VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGPrintCR03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim cr03No_ As String
Dim companyID_ As String
Dim companyName_ As String
Dim issuedDate_ As String
Dim issuedStaffID_ As String
Dim issuedStaffName_ As String
Dim signStaffID_ As String
Dim signStaffName_ As String
Dim signStaffPosition_ As String
Dim isCancel_ As String
Dim colPrintDet_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get CR03No() As String
    CR03No = cr03No_
End Property
Public Property Let CR03No(ByVal newVal As String)
    If Trim(cr03No_) <> Trim(newVal) Then
        cr03No_ = Trim(newVal)
        isChanged_ = True
    End If
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

Public Property Get PrintDetails() As Collection
Dim tmpS As New BWGSearchManager
    If colPrintDet_ Is Nothing Then Set colPrintDet_ = tmpS.CR03DetailSearch("CR03ID='" & id_ & "'")
    Set PrintDetails = colPrintDet_
End Property
Public Property Set PrintDetails(ByVal newVal As Collection)
    Set colPrintDet_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryPrintDetail()
    Set colPrintDet_ = Nothing
End Sub

Public Property Get CompanyName() As String
    CompanyName = companyName_
End Property
Public Property Let CompanyName(ByVal newVal As String)
    If Trim(companyName_) <> Trim(newVal) Then
        companyName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedDate() As String
    IssuedDate = issuedDate_
End Property
Public Property Let IssuedDate(ByVal newVal As String)
    If Trim(issuedDate_) <> Trim(newVal) Then
        issuedDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedStaffID() As String
    IssuedStaffID = issuedStaffID_
End Property
Public Property Let IssuedStaffID(ByVal newVal As String)
    If Trim(issuedStaffID_) <> Trim(newVal) Then
        issuedStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get IssuedStaffName() As String
    IssuedStaffName = issuedStaffName_
End Property
Public Property Let IssuedStaffName(ByVal newVal As String)
    If Trim(issuedStaffName_) <> Trim(newVal) Then
        issuedStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignStaffID() As String
    SignStaffID = signStaffID_
End Property
Public Property Let SignStaffID(ByVal newVal As String)
    If Trim(signStaffID_) <> Trim(newVal) Then
        signStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignStaffName() As String
    SignStaffName = signStaffName_
End Property
Public Property Let SignStaffName(ByVal newVal As String)
    If Trim(signStaffName_) <> Trim(newVal) Then
        signStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignStaffPosition() As String
    SignStaffPosition = signStaffPosition_
End Property
Public Property Let SignStaffPosition(ByVal newVal As String)
    If Trim(signStaffPosition_) <> Trim(newVal) Then
        signStaffPosition_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isCancel() As Boolean
    If isCancel_ = "Y" Then
        isCancel = True
    Else
        isCancel = False
    End If
End Property
Public Property Let isCancel(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isCancel_) <> Trim(tmpStr) Then
        isCancel_ = Trim(tmpStr)
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
Dim x%
    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbPrintCR03Header Where CR03ID='" & id_ & "'", DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("CR03ID") = id_
    End If
    'tmpRec("CR03No") = Trim(cr03No_)
    tmpRec("CompanyID") = Trim(companyID_)
    tmpRec("IssuedDate") = Trim(issuedDate_)
    tmpRec("IssuedStaffID") = Trim(issuedStaffID_)
    tmpRec("SignStaffID") = Trim(signStaffID_)
    If Trim(isCancel_) = "" Then isCancel_ = "N"
    tmpRec("isCancel") = Trim(isCancel_)
    tmpRec.Update
    tmpRec.Requery
    cr03No_ = Trim("" & tmpRec("CR03No"))
    tmpRec.Close
    With DBConn
        .Execute "Delete tbPrintCR03Detail Where CR03ID='" & id_ & "'"
        If Not colPrintDet_ Is Nothing Then
            For x = 1 To colPrintDet_.Count
                colPrintDet_(x).Save
            Next
        End If
    End With
    Set tmpRec = Nothing
    isChanged_ = False
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    With DBConn
        .Execute "Delete tbPrintCR03Detail Where CR03ID='" & id_ & "'"
        .Execute "Delete tbPrintCR03Header Where CR03ID='" & id_ & "'"
    End With
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("CR03ID"))
    cr03No_ = Trim("" & selRec("CR03No"))
    companyID_ = Trim("" & selRec("CompanyID"))
    companyName_ = Trim("" & selRec("CompanyNameTH"))
    issuedDate_ = Trim("" & selRec("IssuedDate"))
    issuedStaffID_ = Trim("" & selRec("IssuedStaffID"))
    issuedStaffName_ = Trim("" & selRec("IssuedStaffName"))
    signStaffID_ = Trim("" & selRec("SignStaffID"))
    signStaffName_ = Trim("" & selRec("SignStaffName"))
    signStaffPosition_ = Trim("" & selRec("PositionName"))
    isCancel_ = Trim("" & selRec("isCancel"))
End Sub

Public Sub CancelBill()
    isCancel_ = "Y"
    DBConnExc "Update tbPrintCR03Header Set isCancel='Y' Where CR03ID='" & id_ & "'"
    isChanged_ = False
End Sub
