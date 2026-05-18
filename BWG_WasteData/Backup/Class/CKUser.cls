VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CKuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim sID_ As String
Dim cID_ As String

Public Property Get sID() As String
    sID = sID_
End Property
Public Property Let sID(ByVal newval As String)
    sID = Trim(newval)
End Property

Public Property Get cID() As String
    cID = cID_
End Property
Public Property Let cID(ByVal newval As String)
    cID = Trim(newval)
End Property

 Public Sub CKusers()
' Dim x%
' Dim tmpRec As ADODB.Recordset
'Set tmpRec = New ADODB.Recordset
'tmpRec.Open "Select CompanyID From tbStaffInCompany Where StaffID = '" & CurrentUser.ID & "' And CompanyID = '776EDA5830F744A6B3551348D851BC22' ", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
'    If tmpRec.EOF = False Then
'       isAkp = False
'    Else
'       isAkp = True
'    End If
 End Sub
Public Sub SCRecordset(selRec As ADODB.Recordset)
    sID_ = Trim("StaffID")
    cID_ = Trim("CompanyID")
End Sub

