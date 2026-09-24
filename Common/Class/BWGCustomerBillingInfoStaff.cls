VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCustomerBillingInfoStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim InFoStaffID_ As String
Dim InFoStaffName_ As String
Dim isChanged_ As Boolean

Public Property Get InFoStaffID() As String
    InFoStaffID = InFoStaffID_
End Property
Public Property Let InFoStaffID(ByVal newVal As String)
    InFoStaffID_ = Trim(newVal)
End Property

Public Property Get InFoStaffName() As String
    InFoStaffName = InFoStaffName_
End Property
Public Property Let InFoStaffName(ByVal newVal As String)
    If Trim(InFoStaffName_) <> Trim(newVal) Then
        InFoStaffName_ = Trim(newVal)
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
    InFoStaffID_ = selRec("InFoStaffID")
    InFoStaffName_ = Trim("" & selRec("InFoStaffName"))
End Sub


