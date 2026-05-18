VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCompanyConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim companyCode_ As String
Dim companyNameTH_ As String
Dim companyNameEN_ As String
Dim shortName_ As String
Dim managerName_ As String
Dim addrTH1_ As String
Dim addrTH2_ As String
Dim addrTH3_ As String
Dim addrEN1_ As String
Dim addrEN2_ As String
Dim addrEN3_ As String
Dim phoneNo_ As String
Dim faxNo_ As String
Dim emailAddr_ As String
Dim website_ As String
Dim companyRem_ As String
Dim taxID_ As String
Dim isDeleted_ As String
Dim deletedDate_ As String
Dim remDesc_ As String
Dim disposerNo_ As String
Dim tsdfID_ As String
Dim tsdfName_ As String
Dim colDisposerNo_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get CompanyCode() As String
    CompanyCode = companyCode_
End Property
Public Property Let CompanyCode(ByVal newVal As String)
    If Trim(companyCode_) <> Trim(newVal) Then
        companyCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyNameTH() As String
    CompanyNameTH = companyNameTH_
End Property
Public Property Let CompanyNameTH(ByVal newVal As String)
    If Trim(companyNameTH_) <> Trim(newVal) Then
        companyNameTH_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyNameEN() As String
    CompanyNameEN = companyNameEN_
End Property
Public Property Let CompanyNameEN(ByVal newVal As String)
    If Trim(companyNameEN_) <> Trim(newVal) Then
        companyNameEN_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ShortName() As String
    ShortName = shortName_
End Property
Public Property Let ShortName(ByVal newVal As String)
    If Trim(shortName_) <> Trim(newVal) Then
        shortName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ManagerName() As String
    ManagerName = managerName_
End Property
Public Property Let ManagerName(ByVal newVal As String)
    If Trim(managerName_) <> Trim(newVal) Then
        managerName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AddrTH1() As String
    AddrTH1 = addrTH1_
End Property
Public Property Let AddrTH1(ByVal newVal As String)
    If Trim(addrTH1_) <> Trim(newVal) Then
        addrTH1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AddrTH2() As String
    AddrTH2 = addrTH2_
End Property
Public Property Let AddrTH2(ByVal newVal As String)
    If Trim(addrTH2_) <> Trim(newVal) Then
        addrTH2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AddrTH3() As String
    AddrTH3 = addrTH3_
End Property
Public Property Let AddrTH3(ByVal newVal As String)
    If Trim(addrTH3_) <> Trim(newVal) Then
        addrTH3_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AddrEN1() As String
    AddrEN1 = addrEN1_
End Property
Public Property Let AddrEN1(ByVal newVal As String)
    If Trim(addrEN1_) <> Trim(newVal) Then
        addrEN1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AddrEN2() As String
    AddrEN2 = addrEN2_
End Property
Public Property Let AddrEN2(ByVal newVal As String)
    If Trim(addrEN2_) <> Trim(newVal) Then
        addrEN2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get AddrEN3() As String
    AddrEN3 = addrEN3_
End Property
Public Property Let AddrEN3(ByVal newVal As String)
    If Trim(addrEN3_) <> Trim(newVal) Then
        addrEN3_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get PhoneNo() As String
    PhoneNo = phoneNo_
End Property
Public Property Let PhoneNo(ByVal newVal As String)
    If Trim(phoneNo_) <> Trim(newVal) Then
        phoneNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get FaxNo() As String
    FaxNo = faxNo_
End Property
Public Property Let FaxNo(ByVal newVal As String)
    If Trim(faxNo_) <> Trim(newVal) Then
        faxNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get EmailAddr() As String
    EmailAddr = emailAddr_
End Property
Public Property Let EmailAddr(ByVal newVal As String)
    If Trim(emailAddr_) <> Trim(newVal) Then
        emailAddr_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get Website() As String
    Website = website_
End Property
Public Property Let Website(ByVal newVal As String)
    If Trim(website_) <> Trim(newVal) Then
        website_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CompanyRemark() As String
    CompanyRemark = companyRem_
End Property
Public Property Let CompanyRemark(ByVal newVal As String)
    If Trim(companyRem_) <> Trim(newVal) Then
        companyRem_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TaxID() As String
    TaxID = taxID_
End Property
Public Property Let TaxID(ByVal newVal As String)
    If Trim(taxID_) <> Trim(newVal) Then
        taxID_ = Trim(newVal)
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

Public Property Get TSDFID() As String
    TSDFID = tsdfID_
End Property
Public Property Let TSDFID(ByVal newVal As String)
    If Trim(tsdfID_) <> Trim(newVal) Then
        tsdfID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TSDFName() As String
    TSDFName = tsdfName_
End Property
Public Property Let TSDFName(ByVal newVal As String)
    If Trim(tsdfName_) <> Trim(newVal) Then
        tsdfName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get DisposerNoList() As Collection
Dim tmpS As New BWGSearchManager
    If colDisposerNo_ Is Nothing Then Set colDisposerNo_ = tmpS.DisposerNoSearch("CompanyID='" & id_ & "'")
    Set DisposerNoList = colDisposerNo_
End Property
Public Property Set DisposerNoList(ByVal newVal As Collection)
    Set colDisposerNo_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryDisposerNo()
    Set colDisposerNo_ = Nothing
End Sub

Public Property Get DeletedDate() As String
   DeletedDate = deletedDate_
End Property
Public Property Let DeletedDate(ByVal newVal As String)
    If Trim(deletedDate_) <> Trim(newVal) Then
        deletedDate_ = Trim(newVal)
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

Public Property Get isDeleted() As Boolean
    If isDeleted_ = "Y" Then
        isDeleted = True
    Else
        isDeleted = False
    End If
End Property
Public Property Let isDeleted(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isDeleted_) <> Trim(tmpStr) Then
        isDeleted_ = Trim(tmpStr)
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
    tmpRec.Open "Select * from tbCompanyConfig Where CompanyID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("CompanyID") = id_
    End If
    tmpRec("CompanyCode") = Trim(companyCode_)
    tmpRec("CompanyNameTH") = Trim(companyNameTH_)
    tmpRec("CompanyNameEN") = Trim(companyNameEN_)
    tmpRec("ShortName") = Trim(shortName_)
    tmpRec("ManagerName") = Trim(managerName_)
    tmpRec("AddrTH1") = Trim(addrTH1_)
    tmpRec("AddrTH2") = Trim(addrTH2_)
    tmpRec("AddrTH3") = Trim(addrTH3_)
    tmpRec("AddrEN1") = Trim(addrEN1_)
    tmpRec("AddrEN2") = Trim(addrEN2_)
    tmpRec("AddrEN3") = Trim(addrEN3_)
    tmpRec("PhoneNo") = Trim(phoneNo_)
    tmpRec("FaxNo") = Trim(faxNo_)
    tmpRec("EmailAddr") = Trim(emailAddr_)
    tmpRec("WebSite") = Trim(website_)
    tmpRec("CompanyRem") = Trim(companyRem_)
    tmpRec("TaxID") = Trim(taxID_)
    tmpRec("isDeleted") = Trim(isDeleted_)
    tmpRec("DeletedDate") = Trim(deletedDate_)
    tmpRec("RemDesc") = Trim(remDesc_)
    tmpRec("DisposerNo") = Trim(disposerNo_)
    tmpRec("TSDFID") = Trim(tsdfID_)
    tmpRec.Update
    tmpRec.Close
    If Not colDisposerNo_ Is Nothing Then
        With DBManager.GetDBConnection
            .Execute "Delete tbDisposerNo Where CompanyID='" & id_ & "'"
            For x = 1 To colDisposerNo_.Count
                If Trim(colDisposerNo_(x).DisposerNo) <> "" Then colDisposerNo_(x).Save
            Next
        End With
    End If
    Set tmpRec = Nothing
    isChanged_ = False
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBManager.GetDBConnection.Execute "Delete tbCompanyConfig Where CompanyID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("CompanyID")
    companyCode_ = Trim("" & selRec("CompanyCode"))
    companyNameTH_ = Trim("" & selRec("CompanyNameTH"))
    companyNameEN_ = Trim("" & selRec("CompanyNameEN"))
    shortName_ = Trim("" & selRec("ShortName"))
    managerName_ = Trim("" & selRec("ManagerName"))
    addrTH1_ = Trim("" & selRec("AddrTH1"))
    addrTH2_ = Trim("" & selRec("AddrTH2"))
    addrTH3_ = Trim("" & selRec("AddrTH3"))
    addrEN1_ = Trim("" & selRec("AddrEN1"))
    addrEN2_ = Trim("" & selRec("AddrEN2"))
    addrEN3_ = Trim("" & selRec("AddrEN3"))
    phoneNo_ = Trim("" & selRec("PhoneNo"))
    faxNo_ = Trim("" & selRec("FaxNo"))
    emailAddr_ = Trim("" & selRec("EmailAddr"))
    website_ = Trim("" & selRec("WebSite"))
    companyRem_ = Trim("" & selRec("CompanyRem"))
    taxID_ = Trim("" & selRec("TaxID"))
    isDeleted_ = Trim("" & selRec("isDeleted"))
    deletedDate_ = Trim("" & selRec("DeletedDate"))
    remDesc_ = Trim("" & selRec("RemDesc"))
    disposerNo_ = Trim("" & selRec("DisposerNo"))
    tsdfID_ = Trim("" & selRec("TSDFID"))
    tsdfName_ = Trim("" & selRec("TSDFName"))
End Sub

