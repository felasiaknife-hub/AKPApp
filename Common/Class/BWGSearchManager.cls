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

Public Function CustomerSearchExtComm(Optional strCond As String = "") As Collection
On Error Resume Next
Dim tmpItem As BWGCustomer
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select  DISTINCT CustomerID, CustomerCode, CustomerName, ProvinceName, SaleStaffFName, SaleStaffLName  from vw_CustomerSearch_ExtComm Order By CustomerName"
    Else
        sqlStr = "Select  DISTINCT CustomerID, CustomerCode, CustomerName, ProvinceName, SaleStaffFName, SaleStaffLName  from vw_CustomerSearch_ExtComm Where (" & strCond & ") Order By CustomerName"
    End If
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomer
        tmpItem.SetDataRS_ExtCom tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerSearchExtComm = tmpCol
    Set tmpCol = Nothing
End Function

Public Function CompanySearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCompanyConfig
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_CompanySearch WHERE isDeleted = 'N' or isDeleted = '' ORDER BY CompanyNameTH"
    Else
        sqlStr = "SELECT * FROM vw_CompanySearch WHERE (" & strCond & ") ORDER BY CompanyNameTH"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCompanyConfig
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CompanySearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function SignNameSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSystemSignPolicies
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbSystemSignPolicies WHERE isActive = 'Y' ORDER BY ConfigID"
    Else
        sqlStr = "SELECT * FROM tbSystemSignPolicies WHERE isActive = 'Y' AND (" & strCond & ") ORDER BY ConfigID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSystemSignPolicies
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "_" & tmpItem.CompanyID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SignNameSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function UserSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCompanyStaff
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_UserSearch WHERE isActive='Y' and isDeleted<>'Y' ORDER BY StaffFName,StaffLName"
    Else
        sqlStr = "SELECT * FROM vw_UserSearch WHERE (" & strCond & ") And isActive='Y' and isDeleted<>'Y' ORDER BY StaffFName,StaffLName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCompanyStaff
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set UserSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function UnitLABSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABUnit
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbLABUnit ORDER BY UnitName"
    Else
        sqlStr = "SELECT * FROM tbLABUnit WHERE (" & strCond & ") ORDER BY UnitName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABUnit
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set UnitLABSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function ReceptionSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGReception
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbReception where  Status='Y' ORDER BY dept,StaffName"
    Else
        sqlStr = "SELECT * FROM tbReception WHERE (" & strCond & " and Status='Y') ORDER BY id"
    End If

    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGReception
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ReceptionSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function UnitGeneralSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGGeneralUnit
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbGeneralUnit ORDER BY UnitName"
    Else
        sqlStr = "SELECT * FROM tbGeneralUnit WHERE (" & strCond & ") ORDER BY UnitName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGGeneralUnit
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set UnitGeneralSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function CancelRemarkSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCancelRemark
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tb_CancelRemark ORDER BY RemID"
    Else
        sqlStr = "SELECT * FROM tb_CancelRemark WHERE (" & strCond & ") ORDER BY RemID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCancelRemark
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.RemID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CancelRemarkSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function ZoneAreaSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGZoneArea
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbZoneArea ORDER BY ZoneName"
    Else
        sqlStr = "SELECT * FROM tbZoneArea WHERE (" & strCond & ") ORDER BY ZoneName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGZoneArea
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ZoneAreaSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function IndustryTypeSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGIndustryType
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbIndustryType ORDER BY TypeName"
    Else
        sqlStr = "SELECT * FROM tbIndustryType WHERE (" & strCond & ") ORDER BY TypeName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGIndustryType
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set IndustryTypeSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function IndustryTypeSearchNew(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGIndustryType_New
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbIndustryType_New ORDER BY Typedesc"
    Else
        sqlStr = "SELECT * FROM tbIndustryType_New WHERE (" & strCond & ") ORDER BY Typedesc"
    End If

    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGIndustryType_New
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.TypeID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set IndustryTypeSearchNew = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function AumphurSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As IMIAumphur
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_AumphurSearch WHERE isDeleted='N' ORDER BY ProvinceName,AumphurName"
    Else
        sqlStr = "SELECT * FROM vw_AumphurSearch WHERE (" & strCond & ") And isDeleted='N' ORDER BY ProvinceName,AumphurName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New IMIAumphur
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set AumphurSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function TumbolSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As IMITumbol
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_TumbolSearch WHERE isDeleted='N' ORDER BY ProvinceName,TumbolName"
    Else
        sqlStr = "SELECT * FROM vw_TumbolSearch WHERE (" & strCond & ") And isDeleted='N' ORDER BY ProvinceName,TumbolName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New IMITumbol
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TumbolSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function ProvinceSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As IMIProvince
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbProvince ORDER BY ProvinceName"
    Else
        sqlStr = "SELECT * FROM tbProvince WHERE (" & strCond & ") ORDER BY ProvinceName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New IMIProvince
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ProvinceSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function CustomerSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustomer
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT *, SubContact = ISNULL((SELECT CompanyName FROM tbCustomer WHERE CustomerID = vw_CustomerSearch.ParentCustID), '') " & vbCrLf
        sqlStr = sqlStr & "FROM vw_CustomerSearch ORDER BY CompanyName"
    Else
        sqlStr = "SELECT *, SubContact = ISNULL((SELECT CompanyName FROM tbCustomer WHERE CustomerID = vw_CustomerSearch.ParentCustID), '')" & vbCrLf
        sqlStr = sqlStr & "FROM vw_CustomerSearch WHERE  (" & strCond & ")  " & vbCrLf
        sqlStr = sqlStr & "ORDER BY CompanyName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomer
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function SaleZoneSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSaleZone
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_SaleZoneSearch  ORDER BY ZoneName"
    Else
        sqlStr = "SELECT * FROM vw_SaleZoneSearch WHERE (" & strCond & ") ORDER BY ZoneName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSaleZone
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SaleZoneSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function EUCode3Search(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGEUCode3
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_EUCode3Search ORDER BY Code3"
    Else
        sqlStr = "SELECT * FROM vw_EUCode3Search WHERE (" & strCond & ") ORDER BY Code3"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGEUCode3
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set EUCode3Search = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function EUCode2Search(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGEUCode2
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_EUCode2Search ORDER BY Code2"
    Else
        sqlStr = "SELECT * FROM vw_EUCode2Search WHERE (" & strCond & ") ORDER BY Code2"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGEUCode2
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set EUCode2Search = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function EUCode1Search(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGEUCode1
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_EUCode1Search ORDER BY Code1"
    Else
        sqlStr = "SELECT * FROM vw_EUCode1Search WHERE (" & strCond & ") ORDER BY Code1"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGEUCode1
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set EUCode1Search = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function WasteDataCRSearch(Optional strCond As String = "", Optional ByVal StrView As String = "vw_WasteDataCRSearch") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGWasteDataCR
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM " & StrView & " ORDER BY WasteNo,WasteName"
    Else
        sqlStr = "SELECT * FROM " & StrView & " WHERE (" & strCond & ") ORDER BY WasteNo,WasteName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGWasteDataCR
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set WasteDataCRSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobAddWasteDataSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGWasteDataCR
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
Dim Dt1 As String
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_WasteDataAddForJob where Status IN ('A', 'IA') ORDER BY QuotationNo,WasteNo,WasteName"
    Else
        sqlStr = "SELECT * FROM vw_WasteDataAddForJob WHERE (" & strCond & ") ORDER BY QuotationNo,WasteNo,WasteName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGWasteDataCR
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "_" & tmpItem.QuotationID & "'"
Err1:
        Set tmpItem = Nothing
'        Debug.Print tmpRec.AbsolutePosition
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobAddWasteDataSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function JobAddWasteDataForTruckSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGWasteDataCR
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
Dim Dt1 As String
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_WasteDataAddForJobFromTruck where Status IN ('A', 'IA') ORDER BY QuotationNo,WasteNo,WasteName"
    Else
        sqlStr = "SELECT * FROM vw_WasteDataAddForJobFromTruck WHERE (" & strCond & ") ORDER BY QuotationNo,WasteNo,WasteName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGWasteDataCR
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "_" & tmpItem.QuotationID & "'"
Err1:
        Set tmpItem = Nothing
'        Debug.Print tmpRec.AbsolutePosition
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobAddWasteDataForTruckSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function JobAddWasteDataSearch2(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGWasteDataCR
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
Dim Dt1 As String
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_WasteDataAddForJob where Status IN ('A', 'IA') ORDER BY QuotationNo,WasteNo,WasteName"
    Else
        sqlStr = "SELECT * FROM vw_WasteDataAddForJob WHERE (" & strCond & ") ORDER BY QuotationNo,WasteNo,WasteName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
    Dt1 = Year(Date) + 541 & "/" & Format(Month(Date), "##") & "/" & Day(Date)
    Dim dt2 As String
    dt2 = Format(Dt1, "YYYY/MM/DD")
    If tmpRec("Approveddate") <= dt2 Then GoTo Err1
        Set tmpItem = New BWGWasteDataCR
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "_" & tmpItem.QuotationID & "'"
Err1:
        Set tmpItem = Nothing
'        Debug.Print tmpRec.AbsolutePosition
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobAddWasteDataSearch2 = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function WasteDataLABSearch(Optional strCond As String = "", Optional isNew_ As Boolean = False) As Collection
On Error GoTo ErrD
Dim tmpItem As BWGWasteDataLAB
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If isNew_ = False Then
        If Trim(strCond) = "" Then
            sqlStr = "SELECT * FROM vw_WasteDataLABSearch"
        Else
            sqlStr = "SELECT * FROM vw_WasteDataLABSearch WHERE (" & strCond & ")"
        End If
    Else
        If Trim(strCond) = "" Then
            sqlStr = "SELECT * FROM vw_WasteDataLAB_NewSearch"
        Else
            sqlStr = "SELECT * FROM vw_WasteDataLAB_NewSearch WHERE (" & strCond & ")"
        End If
    End If
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGWasteDataLAB
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set WasteDataLABSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function TreatmentDataSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTreatmentData
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbTreatmentData ORDER BY TreatmentDesc"
    Else
        sqlStr = "SELECT * FROM tbTreatmentData WHERE (" & strCond & ") ORDER BY TreatmentDesc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTreatmentData
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TreatmentDataSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function WasteTreatmentSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTreatmentData
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_WasteTreatmentSearch ORDER BY TreatmentDesc"
    Else
        sqlStr = "SELECT * FROM vw_WasteTreatmentSearch WHERE (" & strCond & ") ORDER BY TreatmentDesc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTreatmentData
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set WasteTreatmentSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function CustomerSiteSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustomerSite
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_CustomerSiteSearch ORDER BY SiteName"
    Else
        sqlStr = "SELECT * FROM vw_CustomerSiteSearch WHERE (" & strCond & ") ORDER BY SiteName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomerSite
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerSiteSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function CustomerBillingInfoSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustomerBillInfo
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_CustBillingInfoSearch"
    Else
        sqlStr = "SELECT * FROM vw_CustBillingInfoSearch WHERE (" & strCond & ")"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomerBillInfo
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerBillingInfoSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function CustomerBillingAddrSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustomerBillAddr
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_CustomerBillAddrSearch ORDER BY BillName"
    Else
        sqlStr = "SELECT * FROM vw_CustomerBillAddrSearch WHERE (" & strCond & ") ORDER BY BillName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomerBillAddr
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerBillingAddrSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function ContractSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGContract
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_CustomerContractSearch ORDER BY StartDate Desc"
    Else
        sqlStr = "SELECT * FROM vw_CustomerContractSearch WHERE (" & strCond & ") ORDER BY StartDate Desc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGContract
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ContractSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function CmmQuoSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCmmQuoHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_CmmQuoHeader ORDER BY StartDate Desc"
    Else
        sqlStr = "SELECT * FROM vw_CmmQuoHeader WHERE (" & strCond & ") ORDER BY StartDate Desc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCmmQuoHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CmmQuoSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function RemarkListSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGRemarkList
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbRemarkList ORDER BY RemDesc"
    Else
        sqlStr = "SELECT * FROM tbRemarkList WHERE (" & strCond & ") ORDER BY RemDesc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGRemarkList
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set RemarkListSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function QuotationRemark(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGQuotationRem
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbQuotationRemark ORDER BY ItemNo"
    Else
        sqlStr = "SELECT * FROM tbQuotationRemark WHERE (" & strCond & ") ORDER BY ItemNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGQuotationRem
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationRemark = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function TruckTypeSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTruckType
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbTruckType WHERE isActive = 'Y' ORDER BY TruckTypeDesc"
    Else
        sqlStr = "SELECT * FROM tbTruckType WHERE isactive='Y' and (" & strCond & ") ORDER BY TruckTypeDesc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTruckType
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TruckTypeSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function TruckRateSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTruckRate
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_TruckRateSearch WHERE isActive = 'Y' ORDER BY TruckTypeDesc"
    Else
        sqlStr = "SELECT * FROM vw_TruckRateSearch WHERE isactive='Y' and (" & strCond & ") ORDER BY TruckTypeDesc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTruckRate
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TruckRateSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function QuotationTruckRateSearch(Optional strCond As String = "") As Collection

On Error GoTo ErrD
Dim tmpItem As BWGQuotationTruckRate
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_QuotationTruckRateSearch  where isDelete<>'Y' ORDER BY TruckTypeDesc"
    Else
        sqlStr = "SELECT * FROM vw_QuotationTruckRateSearch WHERE  isDelete<>'Y'  and (" & strCond & ") ORDER BY TruckTypeDesc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGQuotationTruckRate
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationTruckRateSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function QuotationTrucSubTypekRateSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGQuotationTruckRate
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
'    If Trim(strCond) = "" Then
'        sqlStr = "SELECT * FROM vw_QuotationTruckRateSearch  where isDelete<>'Y' ORDER BY TruckTypeDesc"
'    Else
'        sqlStr = "SELECT * FROM vw_QuotationTruckRateSearch WHERE  isDelete<>'Y'  and (" & strCond & ") ORDER BY TruckTypeDesc"
'    End If
'
     sqlStr = " select a.* from  vw_QuotationTrucSubTypekRateSearch a join ( "
    sqlStr = sqlStr & " SELECT TruckTypeID,TruckTypeDesc,MAX(Rate) as Rate FROM vw_QuotationTrucSubTypekRateSearch WHERE  isDelete<>'Y'"
    sqlStr = sqlStr & " and (" & strCond & ") group by TruckTypeID,TruckTypeDesc ) as b on"
    sqlStr = sqlStr & " a.TruckTypeID=b.TruckTypeID and a.isDelete<>'Y'and a.Rate=b.Rate and a." & strCond & ""
    sqlStr = sqlStr & " ORDER BY a.TruckTypeDesc"
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGQuotationTruckRate
        If isExist(tmpCol, "'" & tmpRec("SubTypeID") & "'") = False Then
            tmpItem.SetDataFromRecordset2 tmpRec
             tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        End If
       
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationTrucSubTypekRateSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function


Public Function QuotationDetailSearch(Optional strCond As String = "", Optional ByVal StrView As String = "vw_QuotationDetailSearch") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGQuotationDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection

    If StrView <> "vw_QuotationDetailSearch" Then
        If Trim(strCond) = "" Then
            sqlStr = "SELECT * FROM vw_QuotationDetailSearch_New ORDER BY  IssueDate, QuotationNo, RowPos"
        Else
            sqlStr = "SELECT * FROM vw_QuotationDetailSearch_New  WHERE (" & strCond & ") ORDER BY IssueDate, QuotationNo, RowPos"
        End If
        Set tmpRec = New ADODB.Recordset
        Set tmpCol = New Collection
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        
        If tmpRec.RecordCount <= 0 Then
            If Trim(strCond) = "" Then
                sqlStr = "SELECT * FROM " & StrView & " ORDER BY RowPos"
            Else
                sqlStr = "SELECT * FROM " & StrView & " WHERE (" & strCond & ") ORDER BY RowPos"
            End If
            Set tmpRec = New ADODB.Recordset
            Set tmpCol = New Collection
            Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        End If
    
    Else
        If Trim(strCond) = "" Then
            sqlStr = "SELECT * FROM " & StrView & " ORDER BY RowPos"
        Else
            sqlStr = "SELECT * FROM " & StrView & " WHERE (" & strCond & ") ORDER BY RowPos"
        End If
        Set tmpRec = New ADODB.Recordset
        Set tmpCol = New Collection
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    End If
    
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGQuotationDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function QuotationDetailDiscountSearch(Optional strCond As String = "", Optional ByVal StrView As String = "tbQuotationDiscountDetail") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGQTDiscountDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection

    If StrView <> "tbQuotationDiscountDetail" Then
        If Trim(strCond) = "" Then
            sqlStr = "SELECT * FROM tbQuotationDiscountDetail "
        Else
            sqlStr = "SELECT * FROM tbQuotationDiscountDetail  WHERE (" & strCond & ") "
        End If
        Set tmpRec = New ADODB.Recordset
        Set tmpCol = New Collection
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        
        If tmpRec.RecordCount <= 0 Then
            If Trim(strCond) = "" Then
                sqlStr = "SELECT * FROM " & StrView & " "
            Else
                sqlStr = "SELECT * FROM " & StrView & " WHERE (" & strCond & ") "
            End If
            Set tmpRec = New ADODB.Recordset
            Set tmpCol = New Collection
            Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        End If
    
    Else
        If Trim(strCond) = "" Then
            sqlStr = "SELECT * FROM " & StrView & " "
        Else
            sqlStr = "SELECT * FROM " & StrView & " WHERE (" & strCond & ") "
        End If
        Set tmpRec = New ADODB.Recordset
        Set tmpCol = New Collection
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    End If
    
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGQTDiscountDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationDetailDiscountSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function QuotationDetailCmmSearch(Optional strCond As String = "", Optional ByVal StrView As String = "tbQuotationDiscountDetail") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCmmQuoDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection

    If StrView <> "tbCommentQuoDetail" Then
        If Trim(strCond) = "" Then
            sqlStr = "SELECT * FROM tbCommentQuoDetail "
        Else
            sqlStr = "SELECT * FROM tbCommentQuoDetail  WHERE (" & strCond & ") "
        End If
        Set tmpRec = New ADODB.Recordset
        Set tmpCol = New Collection
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        
        If tmpRec.RecordCount <= 0 Then
            If Trim(strCond) = "" Then
                sqlStr = "SELECT * FROM tbCommentQuoDetail "
            Else
                sqlStr = "SELECT * FROM tbCommentQuoDetail WHERE (" & strCond & ") "
            End If
            Set tmpRec = New ADODB.Recordset
            Set tmpCol = New Collection
            Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        End If
    
    Else
        If Trim(strCond) = "" Then
            sqlStr = "SELECT * FROM " & StrView & " "
        Else
            sqlStr = "SELECT * FROM " & StrView & " WHERE (" & strCond & ") "
        End If
        Set tmpRec = New ADODB.Recordset
        Set tmpCol = New Collection
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    End If
    
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCmmQuoDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationDetailCmmSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function


Public Function QuotationHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGQuotationHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_QuotationHeaderSearch ORDER BY QuotationNo"
    Else
        sqlStr = "SELECT * FROM vw_QuotationHeaderSearch WHERE (" & strCond & ")" & vbCrLf & "ORDER BY IssueDate, QuotationNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGQuotationHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function QuotationDiscountHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGQTDiscountHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_QuotationDiscountHeaderSearch ORDER BY QuoDisNo"
    Else
        sqlStr = "SELECT * FROM vw_QuotationDiscountHeaderSearch WHERE (" & strCond & ") ORDER BY QuoDisNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGQTDiscountHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationDiscountHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function QuotationInTripSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGQuotationHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_QuotationInTrip ORDER BY QuotationNo"
    Else
        sqlStr = "SELECT * FROM vw_QuotationInTrip WHERE (" & strCond & ") ORDER BY QuotationNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGQuotationHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationInTripSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function CustomerStaffSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustomerStaff
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbCustomerStaff ORDER BY Fname,LName"
    Else
        sqlStr = "SELECT * FROM tbCustomerStaff WHERE (" & strCond & ") ORDER BY Fname,LName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomerStaff
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerStaffSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function MenifestSearch(Optional strCond As String = "") As Collection
'On Error GoTo ErrD
Dim tmpItem As BWGMenifest
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_MenifestSearch ORDER BY  WorkDate, TimeTableNo, TripNo, CustomerID, MenifestNo, DocNo"
    Else
        sqlStr = "SELECT * FROM vw_MenifestSearch WHERE (" & strCond & ") ORDER BY  WorkDate, TimeTableNo, TripNo, CustomerID, MenifestNo, DocNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGMenifest
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set MenifestSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox err.Description, vbCritical, "Error"
    Exit Function
End Function

Public Function MenifestSearchUnion(Optional strCond As String = "", Optional ByVal strCond2 As String = "") As Collection
'On Error GoTo ErrD
Dim tmpItem As BWGMenifest
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" And Trim(strCond2) = "" Then
        sqlStr = "SELECT * FROM vw_MenifestSearch ORDER BY WorkDate, TimeTableNo, TripNo, CustomerID, MenifestNo, DocNo"
    ElseIf Trim(strCond2) <> "" And Trim(strCond) <> "" Then
        sqlStr = "select distinct * from (SELECT * FROM vw_MenifestSearch WHERE (" & strCond & ")  union   SELECT * FROM vw_MenifestSearch WHERE (" & strCond2 & ") )  as a ORDER BY WorkDate, TimeTableNo, TripNo, CustomerID, MenifestNo, DocNo  "
    Else
         sqlStr = "SELECT * FROM vw_MenifestSearch WHERE (" & strCond & ") ORDER BY WorkDate, TimeTableNo, TripNo, CustomerID, MenifestNo, DocNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGMenifest
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set MenifestSearchUnion = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox err.Description, vbCritical, "Error"
    Exit Function
End Function

Public Function DisposerNoSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGDisposerNo
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbDisposerNo ORDER BY DisposerNo"
    Else
        sqlStr = "SELECT * FROM tbDisposerNo WHERE (" & strCond & ") ORDER BY DisposerNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGDisposerNo
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set DisposerNoSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function KorOrNoSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGKorOrNumber
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbKorOrNumber ORDER BY KorOrNumber"
    Else
        sqlStr = "SELECT * FROM tbKorOrNumber WHERE (" & strCond & ") ORDER BY KorOrNumber"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGKorOrNumber
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set KorOrNoSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_JobDataDetailSearch ORDER BY RowID"
    Else
        sqlStr = "SELECT * FROM vw_JobDataDetailSearch WHERE (" & strCond & ") ORDER BY RowID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDetailInTripSearch(Optional strCond As String = "", Optional SortType As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String, tSql As String, tSql1 As String, tSql2 As String
Dim tmpCol As Collection
    If Trim(SortType) = "" Then
        SortType = "JobNo,RowID"
    End If
    If Trim(strCond) = "" Then
        sqlStr = "(SELECT * FROM vw_JobDetailInTripSearch) Union (SELECT * FROM vw_AddedWasteInTrip) ORDER BY " & SortType
    Else
        tSql1 = "SELECT * FROM vw_JobDetailInTripSearch WHERE (" & strCond & ") ORDER BY " & SortType
        tSql2 = "SELECT * FROM vw_AddedWasteInTrip WHERE (" & strCond & ") ORDER BY " & SortType
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection

    If Trim(strCond) = "" Then
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Function
    Else
        Set tmpRec = GetRS(tSql1, adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Function
    End If
    
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.SetDataFromRecordset tmpRec
        If isExist(tmpCol, "'" & tmpItem.id & "'") = False Then tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set tmpRec = New ADODB.Recordset
    If Trim(strCond) <> "" Then
        Set tmpRec = GetRS(tSql2, adOpenForwardOnly, adLockReadOnly)
        If tmpRec Is Nothing Then Exit Function
    End If
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.SetDataFromRecordset tmpRec
        If isExist(tmpCol, "'" & tmpItem.id & "'") = False Then tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    
    Set JobDetailInTripSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDetailWasteInTripSearch(Optional strCond As String = "", Optional SortType As String = "") As Collection
'On Error GoTo ErrD
Dim tmpItem As BWGJobDataDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String, tSql As String, tSql1 As String, tSql2 As String
Dim tmpCol As Collection
    If Trim(SortType) = "" Then
        SortType = "TimeTableNo, CustomerID, RowID, DocNo, WasteName"
    End If
    If Trim(strCond) = "" Then
        sqlStr = "(SELECT * FROM vw_JobDetailInTripSearch) Union (SELECT * FROM vw_AddedWasteInTrip)"
        tSql = "SELECT * FROM (" & sqlStr & ")" & " AS tmpTripInvoice ORDER BY " & SortType
    Else

        tSql1 = "(SELECT * FROM vw_JobDetailInTripSearch WHERE (" & strCond & ")  AND (ISNULL(IsCanceled,'N') <> 'Y'))"
        tSql2 = "(SELECT * FROM vw_AddedWasteInTrip WHERE (" & strCond & ") AND (ISNULL(IsCanceled,'N') <> 'Y'))"

        tSql = "SELECT * FROM (" & tSql1 & vbCrLf & "UNION" & vbCrLf & tSql2 & ")" & " AS tmpTripInvoice ORDER BY " & SortType
    End If
    Set tmpCol = New Collection
    Set tmpRec = New ADODB.Recordset
    Set tmpRec = GetRS(tSql, adOpenForwardOnly, adLockReadOnly)

    If tmpRec.RecordCount <= 0 Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.SetDataFromRecordset tmpRec
        'If isExist(tmpCol, "'" & tmpItem.ID & "'") = False Then tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        tmpCol.Add tmpItem
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    
    Set JobDetailWasteInTripSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobData
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_JobDataHeaderSearch ORDER BY JobNo"
    Else
        sqlStr = "SELECT * FROM vw_JobDataHeaderSearch WHERE (" & strCond & ")" & vbCrLf & "ORDER BY JobNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobData
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataHeaderBoxSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataBox
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_JobDataHeadeBoxrSearch ORDER BY JobNo"
    Else
        sqlStr = "SELECT * FROM vw_JobDataHeaderBoxSearch WHERE (" & strCond & ") ORDER BY JobNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataBox
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataHeaderBoxSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function CustVisit(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustVisitor
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_CustVisit "
    Else
        sqlStr = "SELECT * FROM vw_CustVisit WHERE (" & strCond & ") "
    End If

    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustVisitor
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustVisit = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataScheduleSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobData
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_JobDataHeaderSchedule ORDER BY JobNo"
    Else
        sqlStr = "SELECT * FROM vw_JobDataHeaderSchedule WHERE (" & strCond & ") ORDER BY JobNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobData
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataScheduleSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataCarSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataCarUsed
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_JobDataCarUsedSearch ORDER BY RowID"
    Else
        sqlStr = "SELECT * FROM vw_JobDataCarUsedSearch WHERE (" & strCond & ") ORDER BY RowID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataCarUsed
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataCarSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataCarBoxSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataCarUsedBox
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_JobDataCarUsedBoxSearch ORDER BY RowID"
    Else
        sqlStr = "SELECT * FROM vw_JobDataCarUsedBoxSearch WHERE (" & strCond & ") ORDER BY RowID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataCarUsedBox
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataCarBoxSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function TruckSubTypeSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTruckSubType
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_TruckSubTypeSearch ORDER BY TruckTypeDesc,SubTypeDesc"
    Else
        sqlStr = "SELECT * FROM vw_TruckSubTypeSearch WHERE (" & strCond & ") ORDER BY TruckTypeDesc,SubTypeDesc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTruckSubType
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TruckSubTypeSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function SupplierContactSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSupplierContact
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
Dim St As String
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbSupplierContact ORDER BY ContactName"
    Else
        sqlStr = "SELECT * FROM tbSupplierContact WHERE (" & strCond & ")  and (stat<> 'n' or Stat is null) ORDER BY ContactName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSupplierContact
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SupplierContactSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function TransportVenderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTransportVendor
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_TransportVenderSearch ORDER BY NameTH"
    Else
        sqlStr = "SELECT * FROM vw_TransportVenderSearch WHERE (" & strCond & ") ORDER BY NameTH"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTransportVendor
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TransportVenderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function VenderCarSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGVenderCar
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_VendorCarSearch ORDER BY CarRegisID"
    Else
        sqlStr = "SELECT * FROM vw_VendorCarSearch WHERE (" & strCond & ") ORDER BY CarRegisID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGVenderCar
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set VenderCarSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataTimeTableTruckSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataTimeTable
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then

'        sqlStr = "SELECT DISTINCT CompanyCode = (SELECT CompanyCode FROM tbCompanyConfig WHERE tbCompanyConfig.CompanyID = vw_JobDataTimeTableTruckSearch.CompanyID), * " & vbCrLf
'        sqlStr = sqlStr & ", ISNULL(CASE WHEN isnull(isTrailer, 'N') = 'Y' THEN" & vbCrLf
'        sqlStr = sqlStr & "                   (SELECT TOP(1) CarRegisID" & vbCrLf
'        sqlStr = sqlStr & "                   FROM          tbVendorCar VD" & vbCrLf
'        sqlStr = sqlStr & "                   WHERE      VD.VendorTruckID IN (TrailTruckID)) ELSE '' END, '') AS TrailerTruckNo" & vbCrLf
'        sqlStr = sqlStr & "FROM vw_JobDataTimeTableTruckSearch " & vbCrLf
'        sqlStr = sqlStr & "ORDER BY WorkDate, TimeTableNo, TripNo"
        sqlStr = " select Distinct CompanyCode,TimetableNo,WorkDate,TruckTypeDesc,CarRegisID,DriverName,FromAumphurName, FromProvinceName,TripNo, " & vbCrLf
        sqlStr = sqlStr & " ToAumphurName,ToProvinceName,VDTransFee  from vw_JobDataTimeTable  " & vbCrLf
        sqlStr = sqlStr & " WHERE and TripNo = 1 and isInvPB <> 'Y' ORDER BY WorkDate, TimeTableNo " & vbCrLf
         
    Else
    
'        sqlStr = "SELECT DISTINCT CompanyCode = (SELECT CompanyCode FROM tbCompanyConfig WHERE tbCompanyConfig.CompanyID = vw_JobDataTimeTableTruckSearch.CompanyID), * " & vbCrLf
'        sqlStr = sqlStr & ", ISNULL(CASE WHEN isnull(isTrailer, 'N') = 'Y' THEN" & vbCrLf
'        sqlStr = sqlStr & "                   (SELECT TOP(1) CarRegisID" & vbCrLf
'        sqlStr = sqlStr & "                   FROM          tbVendorCar VD" & vbCrLf
'        sqlStr = sqlStr & "                   WHERE      VD.VendorTruckID IN (TrailTruckID)) ELSE '' END, '') AS TrailerTruckNo" & vbCrLf
'        sqlStr = sqlStr & "FROM vw_JobDataTimeTableTruckSearch WHERE (" & strCond & ") " & vbCrLf
'''''        sqlStr = sqlStr & "ORDER BY WorkDate, TimeTableNo, TripNo"
        sqlStr = " select Distinct CompanyCode,TimetableNo,WorkDate,TruckTypeDesc,CarRegisID,DriverName,FromAumphurName, FromProvinceName,TripNo, " & vbCrLf
        sqlStr = sqlStr & " ToAumphurName,ToProvinceName,VDTransFee  from vw_JobDataTimeTable  " & vbCrLf
        sqlStr = sqlStr & "  WHERE  (" & strCond & ")  and TripNo = 1 and isInvPB <> 'Y' ORDER BY WorkDate, TimeTableNo " & vbCrLf
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    DoEvents
    Do Until tmpRec.EOF
'        Debug.Print tmpRec.AbsolutePosition
        Set tmpItem = New BWGJobDataTimeTable
        tmpItem.SetDataFromRecordsetBill tmpRec, True
        tmpCol.Add tmpItem, "'" & tmpItem.TimeTableNo & "_" & tmpItem.TripNo & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataTimeTableTruckSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataTimeTableSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataTimeTable
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    
    If Trim(strCond) = "" Then
        sqlStr = "SELECT DISTINCT * " & vbCrLf
        sqlStr = sqlStr & ", ISNULL(CASE WHEN isnull(isTrailer, 'N') = 'Y' THEN" & vbCrLf
        sqlStr = sqlStr & "                   (SELECT TOP(1) CarRegisID" & vbCrLf
        sqlStr = sqlStr & "                   FROM          tbVendorCar VD" & vbCrLf
        sqlStr = sqlStr & "                   WHERE      VD.VendorTruckID IN (TrailTruckID)) ELSE '' END, '') AS TrailerTruckNo" & vbCrLf
        sqlStr = sqlStr & "FROM vw_JobDataTimeTableSearch "
        sqlStr = sqlStr & "ORDER BY WorkDate, TimeTableNo, TripNo"
    Else
        sqlStr = "SELECT DISTINCT * " & vbCrLf
        sqlStr = sqlStr & ", ISNULL(CASE WHEN isnull(isTrailer, 'N') = 'Y' THEN" & vbCrLf
        sqlStr = sqlStr & "                   (SELECT TOP(1) CarRegisID" & vbCrLf
        sqlStr = sqlStr & "                   FROM          tbVendorCar VD" & vbCrLf
        sqlStr = sqlStr & "                   WHERE      VD.VendorTruckID IN (TrailTruckID)) ELSE '' END, '') AS TrailerTruckNo" & vbCrLf
        sqlStr = sqlStr & "FROM vw_JobDataTimeTableSearch WHERE (" & strCond & ") "
        sqlStr = sqlStr & "ORDER BY WorkDate, TimeTableNo, TripNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    DoEvents
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataTimeTable
        tmpItem.SetDataFromRecordset tmpRec
        If isExist(tmpCol, "'" & tmpItem.JobDataCarID & "_" & tmpItem.TripNo & "'") = False Then
            tmpCol.Add tmpItem, "'" & tmpItem.JobDataCarID & "_" & tmpItem.TripNo & "'"
        End If
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataTimeTableSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataTimeTableBoxSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataTimeTableBox
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT DISTINCT * " & vbCrLf
        sqlStr = sqlStr & ", ISNULL(CASE WHEN isnull(isTrailer, 'N') = 'Y' THEN" & vbCrLf
        sqlStr = sqlStr & "                   (SELECT TOP(1) CarRegisID" & vbCrLf
        sqlStr = sqlStr & "                   FROM          tbVendorCar VD" & vbCrLf
        sqlStr = sqlStr & "                   WHERE      VD.VendorTruckID IN (TrailTruckID)) ELSE '' END, '') AS TrailerTruckNo" & vbCrLf
        sqlStr = sqlStr & "FROM vw_JobDataTimeTableBoxSearch "
        sqlStr = sqlStr & "ORDER BY WorkDate, TimeTableNo, TripNo"
    Else
        sqlStr = "SELECT DISTINCT * " & vbCrLf
        sqlStr = sqlStr & ", ISNULL(CASE WHEN isnull(isTrailer, 'N') = 'Y' THEN" & vbCrLf
        sqlStr = sqlStr & "                   (SELECT TOP(1) CarRegisID" & vbCrLf
        sqlStr = sqlStr & "                   FROM          tbVendorCar VD" & vbCrLf
        sqlStr = sqlStr & "                   WHERE      VD.VendorTruckID IN (TrailTruckID)) ELSE '' END, '') AS TrailerTruckNo" & vbCrLf
        sqlStr = sqlStr & "FROM vw_JobDataTimeTableBoxSearch WHERE (" & strCond & ") "
        sqlStr = sqlStr & "ORDER BY WorkDate, TimeTableNo, TripNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    DoEvents
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataTimeTableBox
        tmpItem.SetDataFromRecordset tmpRec
        If isExist(tmpCol, "'" & tmpItem.JobDataCarID & "_" & tmpItem.TripNo & "'") = False Then
            tmpCol.Add tmpItem, "'" & tmpItem.JobDataCarID & "_" & tmpItem.TripNo & "'"
        End If
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataTimeTableBoxSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataTimeTableQuickSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataTimeTable
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_JobDataTimeTableQuickSearch WHERE ISNULL(isCanceled, 'N') <> 'Y' ORDER BY TripNo"
    Else
        sqlStr = "SELECT * FROM vw_JobDataTimeTableQuickSearch WHERE (" & strCond & ") ORDER BY TripNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataTimeTable
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.JobDataCarID & "_" & tmpItem.TripNo & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataTimeTableQuickSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataTimeTableQuickBoxSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataTimeTableBox
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_JobDataTimeTableQuickBoxSearch WHERE ISNULL(isCanceled, 'N') <> 'Y' ORDER BY TripNo"
    Else
        sqlStr = "SELECT * FROM vw_JobDataTimeTableQuickBoxSearch WHERE (" & strCond & ") ORDER BY TripNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataTimeTableBox
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.JobDataCarID & "_" & tmpItem.TripNo & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataTimeTableQuickBoxSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function TripControlInBillSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataTimeTable
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_TripControlInBillSearch ORDER BY WorkDate,TripNo"
    Else
        sqlStr = "SELECT * FROM vw_TripControlInBillSearch WHERE (" & strCond & ") ORDER BY WorkDate,TripNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataTimeTable
        tmpItem.SetDataFromRecordset tmpRec
        If isExist(tmpCol, "'" & tmpItem.JobDataCarID & "_" & tmpItem.TripNo & "'") = False Then
            tmpCol.Add tmpItem, "'" & tmpItem.JobDataCarID & "_" & tmpItem.TripNo & "'"
        End If
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TripControlInBillSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function MethodLABSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABMethodOfAnaly
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbLABMethodOfAnalysis ORDER BY MethodName"
    Else
        sqlStr = "SELECT * FROM tbLABMethodOfAnalysis WHERE (" & strCond & ") ORDER BY MethodName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABMethodOfAnaly
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set MethodLABSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function LABTemplateSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABTemplateDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_LABTemplateDetailSearch ORDER BY ItemNo"
    Else
        sqlStr = "SELECT * FROM vw_LABTemplateDetailSearch WHERE (" & strCond & ") ORDER BY ItemNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABTemplateDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set LABTemplateSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function LABTemplateHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABTemplateHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_LABTemplateHeader ORDER BY TestPriority,TemplateName"
    Else
        sqlStr = "SELECT * FROM vw_LABTemplateHeader WHERE (" & strCond & ") ORDER BY TestPriority,TemplateName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABTemplateHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set LABTemplateHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function ParameterLABSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABParameterList
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_LABParameterList ORDER BY ParameterName"
    Else
        sqlStr = "SELECT * FROM vw_LABParameterList WHERE (" & strCond & ") ORDER BY ParameterName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABParameterList
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ParameterLABSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function TestTypeLABSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABTestType
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbLABTestType ORDER BY TestPriority,TestTypeName"
    Else
        sqlStr = "SELECT * FROM tbLABTestType WHERE (" & strCond & ") ORDER BY TestPriority,TestTypeName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABTestType
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TestTypeLABSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function LabResultDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABResultDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_LabResultDetailSearch ORDER BY ItemNo"
    Else
        sqlStr = "SELECT * FROM vw_LabResultDetailSearch WHERE (" & strCond & ") ORDER BY ItemNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABResultDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set LabResultDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function LabResultHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABResultHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_LabResultHeaderSearch ORDER BY ReportNo"
    Else
        sqlStr = "SELECT * FROM vw_LabResultHeaderSearch WHERE (" & strCond & ") ORDER BY ReportNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABResultHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set LabResultHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function SendSampleDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSendSampleDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_SendSampleDetailSearch ORDER BY RowPos"
    Else
        sqlStr = "SELECT * FROM vw_SendSampleDetailSearch WHERE (" & strCond & ") ORDER BY RowPos"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSendSampleDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SendSampleDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function SendSampleHeadSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSendSampleHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_SendSampleHeaderSearch ORDER BY SendBillNo Desc"
    Else
        sqlStr = "SELECT * FROM vw_SendSampleHeaderSearch WHERE (" & strCond & ") ORDER BY SendBillNo Desc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSendSampleHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SendSampleHeadSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function InternalSampleTypeSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGInternalSampleType
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbInternalSampleType ORDER BY SampleTypeName"
    Else
        sqlStr = "SELECT * FROM tbInternalSampleType WHERE (" & strCond & ") ORDER BY SampleTypeName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGInternalSampleType
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set InternalSampleTypeSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function BilInfoBankAccountSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BwgBankAccountInfo
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_bankaccount "
    Else
        sqlStr = "SELECT * FROM vw_bankaccount WHERE (" & strCond & ") "
    End If

    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BwgBankAccountInfo
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.BankInfo & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set BilInfoBankAccountSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function BilInfoStaffSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGBilllinginfoStaff
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbCustomerBillingInfoStaff ORDER BY InFoStaffID"
    Else
        sqlStr = "SELECT * FROM tbCustomerBillingInfoStaff WHERE (" & strCond & ") ORDER BY InFoStaffID"
    End If

    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGBilllinginfoStaff
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.InFoStaffID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set BilInfoStaffSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function LabFromSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABTemplateHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbLABTemplateHeader  where isclosed is null or isclosed = 'N' or isclosed = ''  ORDER BY TemplateNo"
    Else
        sqlStr = "SELECT * FROM tbLABTemplateHeader WHERE (" & strCond & ") and  (isclosed is null or isclosed = 'N' or isclosed = '')  ORDER BY TemplateNo"
    End If

    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABTemplateHeader
        tmpItem.SetDataFromRecordsetFrom tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set LabFromSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function RemSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGMenifestHeaderRem
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbMenifestHeaderRem ORDER BY RemName"
    Else
        sqlStr = "SELECT * FROM tbMenifestHeaderRem WHERE (" & strCond & ") ORDER BY RemName"
    End If

    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGMenifestHeaderRem
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.RemID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set RemSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function
Public Function CustomerGroupSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustomerGroup
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbCustomerGroup ORDER BY CustGroupName"
    Else
        sqlStr = "SELECT * FROM tbCustomerGroup WHERE (" & strCond & ") ORDER BY CustGroupName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomerGroup
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerGroupSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function SamplingPointSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSamplingPoint
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbSamplingPoint ORDER BY PointCode"
    Else
        sqlStr = "SELECT * FROM tbSamplingPoint WHERE (" & strCond & ") ORDER BY PointCode"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSamplingPoint
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SamplingPointSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function TSDFSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTSDF
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_TSDFSearch ORDER BY CompanyName"
    Else
        sqlStr = "SELECT * FROM vw_TSDFSearch WHERE (" & strCond & ") ORDER BY CompanyName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTSDF
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TSDFSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function FuelBuyRecordSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGFuelBuyRecord
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_FuelBuyRecordSearch ORDER BY BuyDate"
    Else
        sqlStr = "SELECT * FROM vw_FuelBuyRecordSearch WHERE (" & strCond & ") ORDER BY BuyDate"
    End If

    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGFuelBuyRecord
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set FuelBuyRecordSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function VendorExpenseSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGVendorOtherExpense
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_VendorExpenseSearch ORDER BY ExpenseDate"
    Else
        sqlStr = "SELECT * FROM vw_VendorExpenseSearch WHERE (" & strCond & ") ORDER BY ExpenseDate"
    End If

    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGVendorOtherExpense
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set VendorExpenseSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function TransCustChargeRateSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTransVendorChargeRate
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_TransportCustChargeRate"
    Else
        sqlStr = "SELECT * FROM vw_TransportCustChargeRate WHERE (" & strCond & ")"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTransVendorChargeRate
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TransCustChargeRateSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function TransVendorChargeRateSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTransVendorChargeRate
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_TransportVendorChargeRate"
    Else
        sqlStr = "SELECT * FROM vw_TransportVendorChargeRate WHERE (" & strCond & ")"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTransVendorChargeRate
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TransVendorChargeRateSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function SaleDistrictSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSaleDistrict
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_SaleDistrictSearch ORDER BY SDName "
    Else
        sqlStr = "SELECT * FROM vw_SaleDistrictSearch WHERE (" & strCond & ") ORDER BY SDName "
    End If

    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSaleDistrict
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SaleDistrictSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function SaleTargetSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSaleTarget
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_SaleTargetSearch ORDER BY CompanyNameTH,ForYear,ForMonth"
    Else
        sqlStr = "SELECT * FROM vw_SaleTargetSearch WHERE (" & strCond & ") ORDER BY CompanyNameTH,ForYear,ForMonth"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSaleTarget
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SaleTargetSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function ExternalComSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGExternalCom
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_ExternalComSearch ORDER BY WasteName"
    Else
        sqlStr = "SELECT * FROM vw_ExternalComSearch WHERE (" & strCond & ") ORDER BY WasteName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGExternalCom
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ExternalComSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function ComRateConfigSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGComRateConfig
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM ComRateConfig"
    Else
        sqlStr = "SELECT * FROM ComRateConfig WHERE (" & strCond & ")"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGComRateConfig
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ComRateConfigSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function ComRateSaleSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGComRateSale
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM ComRateSale ORDER BY StartPC"
    Else
        sqlStr = "SELECT * FROM ComRateSale WHERE (" & strCond & ") ORDER BY StartPC"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGComRateSale
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ComRateSaleSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function ComRateSupSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGComRateSup
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM ComRateSup ORDER BY StartPC"
    Else
        sqlStr = "SELECT * FROM ComRateSup WHERE (" & strCond & ") ORDER BY StartPC"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGComRateSup
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ComRateSupSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function ComRateDeductSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGComRateDeduct
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM ComRateDeduct ORDER BY StartPC"
    Else
        sqlStr = "SELECT * FROM ComRateDeduct WHERE (" & strCond & ") ORDER BY StartPC"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGComRateDeduct
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ComRateDeductSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function ComRateManagerSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGComRateManager
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM ComRateManager ORDER BY HzStartPC"
    Else
        sqlStr = "SELECT * FROM ComRateManager WHERE (" & strCond & ") ORDER BY HzStartPC"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGComRateManager
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ComRateManagerSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function MenifestDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_MenifestDetailSearch"
    Else
        sqlStr = "SELECT * FROM vw_MenifestDetailSearch WHERE (" & strCond & ")"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set MenifestDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function MenifestDetailAddSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_MenifestDetailSearch ORDER BY WorkDate"
    Else
        sqlStr = "SELECT * FROM vw_MenifestDetailSearch WHERE (" & strCond & ") ORDER BY WorkDate"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "_" & tmpItem.menifestID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set MenifestDetailAddSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function DiscrepWasteSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGDiscrepWaste
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_DiscrepWasteSearch"
    Else
        sqlStr = "SELECT * FROM vw_DiscrepWasteSearch WHERE (" & strCond & ")"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGDiscrepWaste
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set DiscrepWasteSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function BillingDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGBillingDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_BillingDetailSearch ORDER BY RowID"
    Else
        sqlStr = "SELECT * FROM vw_BillingDetailSearch WHERE (" & strCond & ") ORDER BY RowID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGBillingDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set BillingDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function BillingHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGBillingHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_BillHeaderSearch ORDER BY BillingNo"
    Else
        sqlStr = "SELECT * FROM vw_BillHeaderSearch" & vbCrLf & "WHERE (" & strCond & ") " & vbCrLf & "ORDER BY IssuedDate, BillingNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGBillingHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set BillingHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function BillPayHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGBillPayHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_BillPayHeaderSearch ORDER BY BillPayNo"
    Else
        sqlStr = "SELECT * FROM vw_BillPayHeaderSearch WHERE (" & strCond & ") ORDER BY BillPayNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGBillPayHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set BillPayHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function BillPayDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGBillPayDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT DISTINCT * " & vbCrLf
        sqlStr = sqlStr & ", ISNULL(CASE WHEN isnull(isTrailer, 'N') = 'Y' THEN" & vbCrLf
        sqlStr = sqlStr & "                   (SELECT TOP(1) CarRegisID" & vbCrLf
        sqlStr = sqlStr & "                   FROM          tbVendorCar VD" & vbCrLf
        sqlStr = sqlStr & "                   WHERE      VD.VendorTruckID IN (TrailTruckID)) ELSE '' END, '') AS TrailerTruckNo" & vbCrLf
        sqlStr = sqlStr & "FROM vw_BillPayDetailSearch " & vbCrLf
        sqlStr = sqlStr & "ORDER BY RowID"
    Else
        sqlStr = "SELECT DISTINCT * " & vbCrLf
        sqlStr = sqlStr & ", ISNULL(CASE WHEN isnull(isTrailer, 'N') = 'Y' THEN" & vbCrLf
        sqlStr = sqlStr & "                   (SELECT TOP(1) CarRegisID" & vbCrLf
        sqlStr = sqlStr & "                   FROM          tbVendorCar VD" & vbCrLf
        sqlStr = sqlStr & "                   WHERE      VD.VendorTruckID IN (TrailTruckID)) ELSE '' END, '') AS TrailerTruckNo" & vbCrLf
        sqlStr = sqlStr & "FROM vw_BillPayDetailSearch WHERE (" & strCond & ") " & vbCrLf
        sqlStr = sqlStr & "ORDER BY RowID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGBillPayDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set BillPayDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function PaymentHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGPaymentHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_PaymentHeaderSearch ORDER BY PaymentNo"
    Else
        sqlStr = "SELECT * FROM vw_PaymentHeaderSearch WHERE (" & strCond & ") ORDER BY PaymentNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGPaymentHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set PaymentHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function PaymentDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGPaymentDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT DISTINCT * FROM vw_PaymentDetailSearch" & vbCrLf
        sqlStr = sqlStr & "ORDER BY RowID"
    Else
        sqlStr = "SELECT DISTINCT * FROM vw_PaymentDetailSearch WHERE (" & strCond & ") " & vbCrLf
        sqlStr = sqlStr & "ORDER BY RowID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGPaymentDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set PaymentDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function InvCheqPaymentSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGInvCheqPayment
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbInvCheqPayment"
    Else
        sqlStr = "SELECT * FROM tbInvCheqPayment WHERE (" & strCond & ")"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGInvCheqPayment
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set InvCheqPaymentSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function PrintInvDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGPrintInvDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_PrintInvDetSearch ORDER BY RowID"
    Else
        sqlStr = "SELECT * FROM vw_PrintInvDetSearch WHERE (" & strCond & ") ORDER BY RowID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGPrintInvDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set PrintInvDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function PrintInvDetail_AdjustSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGPrintRVDetail_Adjust
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_PrintInvDet_AdjustSearch ORDER BY RowID"
    Else
        sqlStr = "SELECT * FROM vw_PrintInvDet_AdjustSearch WHERE (" & strCond & ") ORDER BY RowID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGPrintRVDetail_Adjust
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set PrintInvDetail_AdjustSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function PrintInvHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGPrintInvHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_PrintInvHeaderSearch ORDER BY InvNo"
    Else
        sqlStr = "SELECT * FROM vw_PrintInvHeaderSearch WHERE (" & strCond & ") ORDER BY InvNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGPrintInvHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set PrintInvHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function WasteGenNoSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGWasteGenNoList
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_WasteGenNoList ORDER BY JobID,WasteCreateNo"
    Else
        sqlStr = "SELECT * FROM vw_WasteGenNoList WHERE (" & strCond & ") ORDER BY JobID,WasteCreateNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGWasteGenNoList
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set WasteGenNoSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function MenifestWasteAddSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGMenifestWasteAdded
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_MenifestWasteAddedSearch ORDER BY RowID"
    Else
        sqlStr = "SELECT * FROM vw_MenifestWasteAddedSearch WHERE (" & strCond & ") ORDER BY RowID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGMenifestWasteAdded
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set MenifestWasteAddSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function CustOrderRemarkSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustOrderRemark
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbCustOrderRemark ORDER BY OrdPos"
    Else
        sqlStr = "SELECT * FROM tbCustOrderRemark WHERE (" & strCond & ") ORDER BY OrdPos"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustOrderRemark
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustOrderRemarkSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function CheqPaymentSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCheqPayment
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM tbCheqPayment ORDER BY PayDate"
    Else
        sqlStr = "SELECT * FROM tbCheqPayment WHERE (" & strCond & ") ORDER BY PayDate"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCheqPayment
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CheqPaymentSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function ApproveLogSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGApproveLog
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_ApproveLogSearch ORDER BY ApproveDate"
    Else
        sqlStr = "SELECT * FROM vw_ApproveLogSearch WHERE (" & strCond & ") ORDER BY ApproveDate"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGApproveLog
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ApproveLogSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function CR03DetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGPrintCR03Det
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_CR03DetailSearch ORDER BY CR03ID,RowPos"
    Else
        sqlStr = "SELECT * FROM vw_CR03DetailSearch WHERE (" & strCond & ") ORDER BY CR03ID,RowPos"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGPrintCR03Det
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CR03DetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function CR03HeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGPrintCR03
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_PrintCR03HeaderSearch ORDER BY CR03No"
    Else
        sqlStr = "SELECT * FROM vw_PrintCR03HeaderSearch WHERE (" & strCond & ") ORDER BY CR03No"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGPrintCR03
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CR03HeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function BillingDetail2Search(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGBillingDetail2
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_BillingDetail2Search ORDER BY RowID"
    Else
        sqlStr = "SELECT * FROM vw_BillingDetail2Search WHERE (" & strCond & ") ORDER BY RowID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGBillingDetail2
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set BillingDetail2Search = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function CR03AddWasteSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_CR03AddWasteSearch ORDER BY WorkDate,DocNo"
    Else
        sqlStr = "SELECT * FROM vw_CR03AddWasteSearch WHERE (" & strCond & ") ORDER BY WorkDate,DocNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.id = Trim("" & tmpRec("JobDetID"))
        tmpItem.menifestID = Trim("" & tmpRec("MenifestID"))
        tmpItem.MenifestWorkDate = Trim("" & tmpRec("WorkDate"))
        tmpItem.MenifestDocNo = Trim("" & tmpRec("DocNo"))
        tmpItem.WasteName = Trim("" & tmpRec("WasteName"))
        tmpItem.WasteEUCode = Trim("" & tmpRec("EUCode"))
        tmpItem.MenifestCarRegisID = Trim("" & tmpRec("CarRegisID"))
        tmpItem.ChargeWeightType = Trim("" & tmpRec("ChargeWType"))
        tmpItem.DisposerWeight = Trim("" & tmpRec("DisposerWeight"))
        tmpItem.EstWasteQty = Trim("" & tmpRec("EstWasteQty"))
        tmpItem.wastedataID = Trim("" & tmpRec("WasteDataID"))
        tmpCol.Add tmpItem, "'" & tmpItem.id & "_" & tmpItem.menifestID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CR03AddWasteSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataTimeTableUpdateSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataTimeTable
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM update_JobDataTimeTableSearch ORDER BY TripNo"
    Else
        sqlStr = "SELECT * FROM update_JobDataTimeTableSearch WHERE (" & strCond & ") ORDER BY TripNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataTimeTable
        tmpItem.SetDataFromRecordset tmpRec
        If isExist(tmpCol, "'" & tmpItem.JobDataCarID & "_" & tmpItem.TripNo & "'") = False Then
            tmpCol.Add tmpItem, "'" & tmpItem.JobDataCarID & "_" & tmpItem.TripNo & "'"
        End If
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataTimeTableUpdateSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function JobDataTimeTableUpdateSearchTmp(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataTimeTable
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select top 1000 * from update_JobDataTimeTableSearch ORDER BY TripNo"
    Else
        sqlStr = "Select top 1000 * from update_JobDataTimeTableSearch WHERE (" & strCond & ") ORDER BY TripNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataTimeTable
        tmpItem.SetDataFromRecordset tmpRec
        If isExist(tmpCol, "'" & tmpItem.JobDataCarID & "_" & tmpItem.TripNo & "'") = False Then
            tmpCol.Add tmpItem, "'" & tmpItem.JobDataCarID & "_" & tmpItem.TripNo & "'"
        End If
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataTimeTableUpdateSearchTmp = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function AdjustDeptHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGAdjDeptHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_AdjDeptHeaderSearch ORDER BY AdjustBillNo"
    Else
        sqlStr = "SELECT * FROM vw_AdjDeptHeaderSearch WHERE (" & strCond & ") ORDER BY AdjustBillNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGAdjDeptHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set AdjustDeptHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function AdjustDeptDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGAdjDeptDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_AdjDeptDetSearch ORDER BY RowPos"
    Else
        sqlStr = "SELECT * FROM vw_AdjDeptDetSearch WHERE (" & strCond & ") ORDER BY RowPos"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGAdjDeptDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set AdjustDeptDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg
    Exit Function
End Function

Public Function AdjustDeptItemSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGAdjDeptDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "SELECT * FROM vw_BillInvItemSearch ORDER BY BillingNo,DocNo"
    Else
        sqlStr = "SELECT * FROM vw_BillInvItemSearch WHERE (" & strCond & ") ORDER BY BillingNo,DocNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec Is Nothing Then Exit Function
    Do Until tmpRec.EOF
        Set tmpItem = New BWGAdjDeptDetail
        tmpItem.AdjBillDesc = Trim("" & tmpRec("WasteName"))
        
        If Trim("" & tmpRec("ChargeWType")) = "A" Then
            tmpItem.AdjBillQty = tmpRec("DisposerWeight")
        Else
            tmpItem.AdjBillQty = tmpRec("CustWeight")
        End If
        If InStr(1, Trim("" & tmpRec("TreatmentUnitID")), "เที่ยว") > 0 Then
            tmpItem.AdjUnitName = "เที่ยว"
        ElseIf InStr(1, Trim("" & tmpRec("TreatmentUnitID")), "ตัน") > 0 Then
            tmpItem.AdjUnitName = "ตัน"
        Else
            tmpItem.AdjUnitName = ""
        End If
        tmpItem.AdjUnitPrice = tmpRec("TreatmentRate")
        tmpItem.BillingHeaderID = Trim("" & tmpRec("BillingHeaderID"))
        tmpItem.BillingNo = Trim("" & tmpRec("BillingNo"))
        tmpItem.JobDetID = Trim("" & tmpRec("JobDetID"))
        tmpItem.ManifestID = Trim("" & tmpRec("MenifestID"))
        tmpItem.MenifestDocNo = Trim("" & tmpRec("DocNo"))
        tmpItem.MenifestNo = Trim("" & tmpRec("MenifestNo"))
        tmpItem.id = GetGUID
        tmpCol.Add tmpItem, "'" & tmpItem.BillingHeaderID & "_" & tmpItem.ManifestID & "_" & tmpItem.JobDetID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set AdjustDeptItemSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    GetMsg "", vbCritical
    Exit Function
End Function

Public Function MenifestSearchRE(Optional strCond As String = "") As Collection
On Error Resume Next
Dim tmpItem As BWGMenifestBillRE
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
     sqlStr = "Select distinct * from vw_Me_Bill_RE WHERE (" & strCond & ") ORDER BY WorkDate, TimeTableNo, DocNo"
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGMenifestBillRE
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.id & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set MenifestSearchRE = tmpCol
    Set tmpCol = Nothing
End Function

