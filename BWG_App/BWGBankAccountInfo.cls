VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BwgBankAccountInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Dim BankCode_ As String
Dim AccountNo_ As String
Dim AccountName_ As String
Dim BankInfo_ As String
Dim isChanged_ As Boolean

Public Property Get BankCode() As String
    BankCode = BankCode_
End Property
Public Property Let BankCode(ByVal newVal As String)
    BankCode_ = Trim(newVal)
End Property

Public Property Get AccountNo() As String
    AccountNo = AccountNo_
End Property
Public Property Let AccountNo(ByVal newVal As String)
    If Trim(AccountNo_) <> Trim(newVal) Then
        AccountNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get AccountName() As String
    AccountName = AccountName_
End Property
Public Property Let AccountName(ByVal newVal As String)
    If Trim(AccountName_) <> Trim(newVal) Then
        AccountName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property
Public Property Get BankInfo() As String
    BankInfo = BankInfo_
End Property
Public Property Let BankInfo(ByVal newVal As String)
    If Trim(BankInfo_) <> Trim(newVal) Then
        BankInfo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    BankCode_ = selRec("BankCode")
    AccountNo_ = Trim("" & selRec("AccountNo"))
    AccountName_ = Trim("" & selRec("AccountName"))
    BankInfo_ = Trim("" & selRec("BankInfo"))
End Sub

   

