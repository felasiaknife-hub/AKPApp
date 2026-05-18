VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGSearchManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function SaleZoneSearch(Optional strCond As String = "") As Collection
Dim tmpItem As BWGSaleZone
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbSaleZone Order By ZoneName"
    Else
        sqlStr = "Select * from tbSaleZone Where (" & strCond & ") Order By ZoneName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSaleZone
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SaleZoneSearch = tmpCol
    Set tmpCol = Nothing
End Function

Public Function UserSearch(Optional strCond As String = "") As Collection
Dim tmpItem As BWGCompanyStaff
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_UserSearch Where isActive='Y' and isDeleted<>'Y' Order By StaffFName,StaffLName"
    Else
        sqlStr = "Select * from vw_UserSearch Where (" & strCond & ") And isActive='Y' and isDeleted<>'Y' Order By StaffFName,StaffLName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCompanyStaff
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set UserSearch = tmpCol
    Set tmpCol = Nothing
End Function

Public Function WasteTreatmentSearch(Optional strCond As String = "") As Collection
Dim tmpItem As BWGTreatmentData
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_WasteTreatmentSearch Order By TreatmentDesc"
    Else
        sqlStr = "Select * from vw_WasteTreatmentSearch Where (" & strCond & ") Order By TreatmentDesc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTreatmentData
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set WasteTreatmentSearch = tmpCol
    Set tmpCol = Nothing
End Function

Public Function CR03DetailSearch(Optional strCond As String = "") As Collection
Dim tmpItem As BWGPrintCR03Det
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_CR03DetailSearch Order By CR03ID,RowPos"
    Else
        sqlStr = "Select * from vw_CR03DetailSearch Where (" & strCond & ") Order By CR03ID,RowPos"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGPrintCR03Det
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CR03DetailSearch = tmpCol
    Set tmpCol = Nothing
End Function

Public Function MenifestDetailAddSearch(Optional strCond As String = "") As Collection
Dim tmpItem As BWGJobDataDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_MenifestDetailSearch Order By WorkDate"
    Else
        sqlStr = "Select * from vw_MenifestDetailSearch Where (" & strCond & ") Order By WorkDate"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "_" & tmpItem.MenifestID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set MenifestDetailAddSearch = tmpCol
    Set tmpCol = Nothing
End Function

Public Function InvSearch(Optional strCond As String = "") As Collection
Dim tmpItem As BWGInvData
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from Com_vwInvoiceSearch Order By InvDate"
    Else
        sqlStr = "Select * from Com_vwInvoiceSearch Where (" & strCond & ") Order By InvDate"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGInvData
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.InvID & "_" & tmpItem.MenifestID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set InvSearch = tmpCol
    Set tmpCol = Nothing
End Function

Public Function DisposerNoSearch(Optional strCond As String = "") As Collection
Dim tmpItem As BWGDisposerNo
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbDisposerNo Order By DisposerNo"
    Else
        sqlStr = "Select * from tbDisposerNo Where (" & strCond & ") Order by DisposerNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGDisposerNo
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set DisposerNoSearch = tmpCol
    Set tmpCol = Nothing
End Function

Public Function CustomerSearch(Optional strCond As String = "") As Collection
On Error Resume Next
Dim tmpItem As BWGCustomer
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_CustomerSearch Order By CompanyName"
    Else
        sqlStr = "Select * from vw_CustomerSearch Where (" & strCond & ") Order By CompanyName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
'    tmpRec.Open sqlStr, DBConn, adOpenForwardOnly, adLockReadOnly
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomer
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerSearch = tmpCol
    Set tmpCol = Nothing
End Function
Public Function MenifestSearchRE(Optional strCond As String = "") As Collection
On Error Resume Next
Dim tmpItem As BWGMenifestBillRE
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
     sqlStr = "Select distinct * from vw_Me_Bill_RE Where (" & strCond & ") Order By WorkDate"
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGMenifestBillRE
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set MenifestSearchRE = tmpCol
    Set tmpCol = Nothing
End Function

Public Function CustomerSearchExtComm(Optional strCond As String = "") As Collection
On Error Resume Next
Dim tmpItem As BWGCustomer
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select  DISTINCT CustomerID, CustomerCode, CompanyName, ProvinceName, SaleStaffFName, SaleStaffLName  from vw_CustomerSearch_ExtComm Order By CompanyName"
    Else
        sqlStr = "Select  DISTINCT CustomerID, CustomerCode, CompanyName, ProvinceName, SaleStaffFName, SaleStaffLName  from vw_CustomerSearch_ExtComm Where (" & strCond & ") Order By CompanyName"
    End If
        Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomer
        tmpItem.SetDataRS_ExtCom tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerSearchExtComm = tmpCol
    Set tmpCol = Nothing
End Function

Public Function CR03HeaderSearch(Optional strCond As String = "") As Collection
Dim tmpItem As BWGPrintCR03
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_PrintCR03HeaderSearch Order By CR03No"
    Else
        sqlStr = "Select * from vw_PrintCR03HeaderSearch Where (" & strCond & ") Order By CR03No"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGPrintCR03
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CR03HeaderSearch = tmpCol
    Set tmpCol = Nothing
End Function

Public Function TruckTypeSearch(Optional strCond As String = "") As Collection
Dim tmpItem As BWGTruckType
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbTruckType Order By TruckTypeDesc"
    Else
        sqlStr = "Select * from tbTruckType Where (" & strCond & ") Order by TruckTypeDesc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTruckType
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TruckTypeSearch = tmpCol
    Set tmpCol = Nothing
End Function

Public Function CompanySearch(Optional strCond As String = "") As Collection
Dim tmpItem As BWGCompanyConfig
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_CompanySearch Where isDeleted = 'N' or isDeleted = '' Order By CompanyNameTH"
    Else
        sqlStr = "Select * from vw_CompanySearch Where (" & strCond & ") Order By CompanyNameTH"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCompanyConfig
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CompanySearch = tmpCol
    Set tmpCol = Nothing
End Function

