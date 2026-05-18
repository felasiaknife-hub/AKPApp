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

Public Function CompanySearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
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
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
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
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function UserSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
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
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
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
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function UnitLABSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABUnit
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbLABUnit Order By UnitName"
    Else
        sqlStr = "Select * from tbLABUnit Where (" & strCond & ") Order By UnitName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABUnit
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set UnitLABSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function UnitGeneralSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGGeneralUnit
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbGeneralUnit Order By UnitName"
    Else
        sqlStr = "Select * from tbGeneralUnit Where (" & strCond & ") Order By UnitName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGGeneralUnit
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set UnitGeneralSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function ZoneAreaSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGZoneArea
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbZoneArea Order By ZoneName"
    Else
        sqlStr = "Select * from tbZoneArea Where (" & strCond & ") Order By ZoneName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGZoneArea
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ZoneAreaSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function IndustryTypeSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGIndustryType
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbIndustryType Order By TypeName"
    Else
        sqlStr = "Select * from tbIndustryType Where (" & strCond & ") Order By TypeName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGIndustryType
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set IndustryTypeSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function AumphurSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As IMIAumphur
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_AumphurSearch Where isDeleted='N' Order By ProvinceName,AumphurName"
    Else
        sqlStr = "Select * from vw_AumphurSearch Where (" & strCond & ") And isDeleted='N' Order By ProvinceName,AumphurName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New IMIAumphur
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set AumphurSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function TumbolSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As IMITumbol
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_TumbolSearch Where isDeleted='N' Order By ProvinceName,TumbolName"
    Else
        sqlStr = "Select * from vw_TumbolSearch Where (" & strCond & ") And isDeleted='N' Order By ProvinceName,TumbolName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New IMITumbol
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TumbolSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function ProvinceSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As IMIProvince
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbProvince Order By ProvinceName"
    Else
        sqlStr = "Select * from tbProvince Where (" & strCond & ") Order By ProvinceName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New IMIProvince
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ProvinceSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function CustomerSearch(Optional strCond As String = "", Optional ByVal isNew_ As Boolean = False) As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustomer
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String, Cond As String
Dim tmpCol As Collection, i As Integer

    Set tmpCol = New Collection
    For i = 1 To 2
        If i = 1 Then
            If Trim(strCond) = "" Then
                sqlStr = "Select * from vw_CustomerSearch Order By CompanyName"
            Else
                sqlStr = "Select * from vw_CustomerSearch Where (" & strCond & ") Order By CompanyName"
            End If
        Else
            If Trim(strCond) = "" Then
                sqlStr = "Select * from vw_CustomerNew_Search Order By CompanyName"
            Else
                Cond = strCond & " AND (isnull(isNew, '') = 'Y')"
                sqlStr = "Select * from vw_CustomerNew_Search Where (" & Cond & ") Order By CompanyName"
            End If
        End If
    
        Set tmpRec = New ADODB.Recordset
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'        tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
        Do Until tmpRec.EOF
            Set tmpItem = New BWGCustomer
            tmpItem.SetDataFromRecordset tmpRec
            tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
            Set tmpItem = Nothing
            tmpRec.MoveNext
        Loop
        tmpRec.Close
        Set tmpRec = Nothing
    Next
    
    Set CustomerSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function SaleZoneSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
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
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
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
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function EUCode3Search(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGEUCode3
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_EUCode3Search Order By Code3"
    Else
        sqlStr = "Select * from vw_EUCode3Search Where (" & strCond & ") Order By Code3"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGEUCode3
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set EUCode3Search = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function EUCode2Search(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGEUCode2
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_EUCode2Search Order By Code2"
    Else
        sqlStr = "Select * from vw_EUCode2Search Where (" & strCond & ") Order By Code2"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGEUCode2
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set EUCode2Search = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function EUCode1Search(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGEUCode1
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_EUCode1Search Order By Code1"
    Else
        sqlStr = "Select * from vw_EUCode1Search Where (" & strCond & ") Order By Code1"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGEUCode1
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set EUCode1Search = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function WasteDataCRSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGWasteDataCR
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String, Str1 As String
Dim tmpCol As Collection
Dim i As Integer
    Set tmpCol = New Collection
    For i = 1 To 2
        If i = 1 Then
            If Trim(strCond) = "" Then
                sqlStr = "Select * from vw_WasteDataCRSearch Order By WasteNo,WasteName"
            Else
                sqlStr = "Select * from vw_WasteDataCRSearch Where (" & strCond & ") Order By WasteNo,WasteName"
            End If
        ElseIf i = 2 Then
            If Trim(strCond) = "" Then
                sqlStr = "Select * from vw_WasteDataCR_NewSearch Order By WasteNo,WasteName"
            Else
                Str1 = strCond & " AND (isNew = 'Y')"
                sqlStr = "Select * from vw_WasteDataCR_NewSearch Where (" & Str1 & ") Order By WasteNo,WasteName"
            End If
        End If
        
        Set tmpRec = New ADODB.Recordset
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'        tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
        Do Until tmpRec.EOF
            Set tmpItem = New BWGWasteDataCR
            tmpItem.SetDataFromRecordset tmpRec
            tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
            Set tmpItem = Nothing
            tmpRec.MoveNext
        Loop
        tmpRec.Close
        Set tmpRec = Nothing
    Next i
    
    Set WasteDataCRSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function JobAddWasteDataSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGWasteDataCR
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_WasteDataAddForJob Order By QuotationNo,WasteNo,WasteName"
    Else
        sqlStr = "Select * from vw_WasteDataAddForJob Where (" & strCond & ") Order By QuotationNo,WasteNo,WasteName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGWasteDataCR
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "_" & tmpItem.QuotationID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobAddWasteDataSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
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
            sqlStr = "Select * from vw_WasteDataLABSearch"
        Else
            sqlStr = "Select * from vw_WasteDataLABSearch Where (" & strCond & ")"
        End If
    Else
        If Trim(strCond) = "" Then
            sqlStr = "Select * from vw_WasteDataLAB_NewSearch"
        Else
            sqlStr = "Select * from vw_WasteDataLAB_NewSearch Where (" & strCond & ")"
        End If
    End If
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGWasteDataLAB
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set WasteDataLABSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function TreatmentDataSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTreatmentData
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbTreatmentData Order By TreatmentDesc"
    Else
        sqlStr = "Select * from tbTreatmentData Where (" & strCond & ") Order By TreatmentDesc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTreatmentData
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TreatmentDataSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function WasteTreatmentSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
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
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
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
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function CustomerSiteSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustomerSite
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_CustomerSiteSearch Order By SiteName"
    Else
        sqlStr = "Select * from vw_CustomerSiteSearch Where (" & strCond & ") Order By SiteName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomerSite
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerSiteSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function CustomerBillingInfoSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustomerBillInfo
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_CustBillingInfoSearch"
    Else
        sqlStr = "Select * from vw_CustBillingInfoSearch Where (" & strCond & ")"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomerBillInfo
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerBillingInfoSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function CustomerBillingAddrSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustomerBillAddr
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_CustomerBillAddrSearch Order By BillName"
    Else
        sqlStr = "Select * from vw_CustomerBillAddrSearch Where (" & strCond & ") Order by BillName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomerBillAddr
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerBillingAddrSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function ContractSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGContract
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_CustomerContractSearch Order By StartDate Desc"
    Else
        sqlStr = "Select * from vw_CustomerContractSearch Where (" & strCond & ") Order by StartDate Desc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGContract
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ContractSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function RemarkListSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGRemarkList
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbRemarkList Order By RemDesc"
    Else
        sqlStr = "Select * from tbRemarkList Where (" & strCond & ") Order by RemDesc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGRemarkList
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set RemarkListSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function QuotationRemark(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGQuotationRem
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbQuotationRemark Order By ItemNo"
    Else
        sqlStr = "Select * from tbQuotationRemark Where (" & strCond & ") Order by ItemNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGQuotationRem
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationRemark = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function TruckTypeSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
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
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
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
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function QuotationDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGQuotationDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_QuotationDetailSearch Order By RowPos"
    Else
        sqlStr = "Select * from vw_QuotationDetailSearch Where (" & strCond & ") Order by RowPos"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGQuotationDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function QuotationHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGQuotationHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_QuotationHeaderSearch Order By QuotationNo"
    Else
        sqlStr = "Select * from vw_QuotationHeaderSearch Where (" & strCond & ") Order by QuotationNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGQuotationHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function QuotationInTripSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGQuotationHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_QuotationInTrip Order By QuotationNo"
    Else
        sqlStr = "Select * from vw_QuotationInTrip Where (" & strCond & ") Order by QuotationNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGQuotationHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set QuotationInTripSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function CustomerStaffSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustomerStaff
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbCustomerStaff Order By Fname,LName"
    Else
        sqlStr = "Select * from tbCustomerStaff Where (" & strCond & ") Order by Fname,LName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomerStaff
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerStaffSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function



Public Function DisposerNoSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
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
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
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
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function KorOrNoSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGKorOrNumber
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbKorOrNumber Order By KorOrNumber"
    Else
        sqlStr = "Select * from tbKorOrNumber Where (" & strCond & ") Order by KorOrNumber"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGKorOrNumber
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set KorOrNoSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function JobDataDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_JobDataDetailSearch Order By RowID"
    Else
        sqlStr = "Select * from vw_JobDataDetailSearch Where (" & strCond & ") Order by RowID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set JobDataDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
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
        sqlStr = "(Select * from vw_JobDetailInTripSearch) Union (Select * from vw_AddedWasteInTrip) Order By " & SortType
    Else
        tSql1 = "Select * from vw_JobDetailInTripSearch Where (" & strCond & ") Order by " & SortType
        tSql2 = "Select * from vw_AddedWasteInTrip Where (" & strCond & ") Order by " & SortType
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    
'    tSql = "ALTER VIEW [dbo].[vw_JobDetailWasteInTripSearch]" & vbCrLf
'    tSql = tSql & "AS" & vbCrLf
'    tSql = tSql & sqlStr
'    DBManager.GetDBConnection.Execute tSql
'    DoEvents
'    tmpRec.Open "Exec SP_JobDetailWasteInTripSearch", DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    If Trim(strCond) = "" Then
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'        tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Else
        Set tmpRec = GetRS(tSql1, adOpenForwardOnly, adLockReadOnly)
'        tmpRec.Open tSql1, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    End If
    
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.SetDataFromRecordset tmpRec
        If isExist(tmpCol, "'" & tmpItem.ID & "'") = False Then tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set tmpRec = New ADODB.Recordset
    If Trim(strCond) <> "" Then
        Set tmpRec = GetRS(tSql2, adOpenForwardOnly, adLockReadOnly)
'        tmpRec.Open tSql2, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    End If
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.SetDataFromRecordset tmpRec
        If isExist(tmpCol, "'" & tmpItem.ID & "'") = False Then tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    
    Set JobDetailInTripSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function JobDetailWasteInTripSearch(Optional strCond As String = "", Optional SortType As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String, tSql As String, tSql1 As String, tSql2 As String
Dim tmpCol As Collection
    If Trim(SortType) = "" Then
        SortType = "JobNo,RowID"
    End If
    If Trim(strCond) = "" Then
        sqlStr = "(Select * from vw_JobDetailInTripSearch) Union (Select * from vw_AddedWasteInTrip) Order By " & SortType
    Else
        tSql1 = "(Select * from vw_JobDetailInTripSearch Where (" & strCond & ")) Order by " & SortType
        tSql2 = "(Select * from vw_AddedWasteInTrip Where (" & strCond & ")) Order by " & SortType
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    
'    tSql = "ALTER VIEW [dbo].[vw_JobDetailWasteInTripSearch]" & vbCrLf
'    tSql = tSql & "AS" & vbCrLf
'    tSql = tSql & sqlStr
'    DBManager.GetDBConnection.Execute tSql
'    tmpRec.Open "Exec SP_JobDetailWasteInTripSearch", DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    If Trim(strCond) = "" Then
        Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'        tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Else
        Set tmpRec = GetRS(tSql1, adOpenForwardOnly, adLockReadOnly)
'        tmpRec.Open tSql1, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    End If
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
    
    Set tmpRec = New ADODB.Recordset
    If Trim(strCond) <> "" Then
        Set tmpRec = GetRS(tSql2, adOpenForwardOnly, adLockReadOnly)
'        tmpRec.Open tSql2, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    End If
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
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function TruckSubTypeSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTruckSubType
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_TruckSubTypeSearch Order By TruckTypeDesc,SubTypeDesc"
    Else
        sqlStr = "Select * from vw_TruckSubTypeSearch Where (" & strCond & ") Order by TruckTypeDesc,SubTypeDesc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTruckSubType
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TruckSubTypeSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function SupplierContactSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSupplierContact
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbSupplierContact Order By ContactName"
    Else
        sqlStr = "Select * from tbSupplierContact Where (" & strCond & ") Order by ContactName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSupplierContact
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SupplierContactSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function TransportVenderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTransportVendor
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_TransportVenderSearch Order By NameTH"
    Else
        sqlStr = "Select * from vw_TransportVenderSearch Where (" & strCond & ") Order by NameTH"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTransportVendor
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TransportVenderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function VenderCarSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGVenderCar
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_VendorCarSearch Order By CarRegisID"
    Else
        sqlStr = "Select * from vw_VendorCarSearch Where (" & strCond & ") Order by CarRegisID"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGVenderCar
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set VenderCarSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function MethodLABSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABMethodOfAnaly
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbLABMethodOfAnalysis Order By MethodName"
    Else
        sqlStr = "Select * from tbLABMethodOfAnalysis Where (" & strCond & ") Order By MethodName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABMethodOfAnaly
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set MethodLABSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function LABTemplateSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABTemplateDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_LABTemplateDetailSearch Order By ItemNo"
    Else
        sqlStr = "Select * from vw_LABTemplateDetailSearch Where (" & strCond & ") Order by ItemNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABTemplateDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set LABTemplateSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function LABTemplateHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABTemplateHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_LABTemplateHeader Order By TestPriority,TemplateName"
    Else
        sqlStr = "Select * from vw_LABTemplateHeader Where (" & strCond & ") Order by TestPriority,TemplateName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABTemplateHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set LABTemplateHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function ParameterLABSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABParameterList
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_LABParameterList Order By ParameterName"
    Else
        sqlStr = "Select * from vw_LABParameterList Where (" & strCond & ") Order By ParameterName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABParameterList
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set ParameterLABSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function TestTypeLABSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABTestType
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbLABTestType Order By TestPriority,TestTypeName"
    Else
        sqlStr = "Select * from tbLABTestType Where (" & strCond & ") Order By TestPriority,TestTypeName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABTestType
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TestTypeLABSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function LabResultDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABResultDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_LabResultDetailSearch Order By ItemNo"
    Else
        sqlStr = "Select * from vw_LabResultDetailSearch Where (" & strCond & ") Order By ItemNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABResultDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set LabResultDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function LabResultHeaderSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGLABResultHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_LabResultHeaderSearch Order By ReportNo"
    Else
        sqlStr = "Select * from vw_LabResultHeaderSearch Where (" & strCond & ") Order By ReportNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGLABResultHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set LabResultHeaderSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function SendSampleDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSendSampleDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_SendSampleDetailSearch Order By RowPos"
    Else
        sqlStr = "Select * from vw_SendSampleDetailSearch Where (" & strCond & ") Order By RowPos"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSendSampleDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SendSampleDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function SendSampleHeadSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSendSampleHeader
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_SendSampleHeaderSearch Order By SendBillNo Desc"
    Else
        sqlStr = "Select * from vw_SendSampleHeaderSearch Where (" & strCond & ") Order By SendBillNo Desc"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSendSampleHeader
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SendSampleHeadSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function InternalSampleTypeSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGInternalSampleType
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbInternalSampleType Order By SampleTypeName"
    Else
        sqlStr = "Select * from tbInternalSampleType Where (" & strCond & ") Order By SampleTypeName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGInternalSampleType
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set InternalSampleTypeSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function CustomerGroupSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustomerGroup
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbCustomerGroup Order By CustGroupName"
    Else
        sqlStr = "Select * from tbCustomerGroup Where (" & strCond & ") Order By CustGroupName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustomerGroup
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustomerGroupSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function SamplingPointSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSamplingPoint
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbSamplingPoint Order By PointCode"
    Else
        sqlStr = "Select * from tbSamplingPoint Where (" & strCond & ") Order By PointCode"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSamplingPoint
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SamplingPointSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function TSDFSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGTSDF
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_TSDFSearch Order By CompanyName"
    Else
        sqlStr = "Select * from vw_TSDFSearch Where (" & strCond & ") Order By CompanyName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGTSDF
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set TSDFSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function







Public Function SaleDistrictSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGSaleDistrict
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_SaleDistrictSearch Order By SDName"
    Else
        sqlStr = "Select * from vw_SaleDistrictSearch Where (" & strCond & ") Order By SDName"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGSaleDistrict
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set SaleDistrictSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function














Public Function MenifestDetailSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_MenifestDetailSearch"
    Else
        sqlStr = "Select * from vw_MenifestDetailSearch Where (" & strCond & ")"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set MenifestDetailSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function MenifestDetailAddSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
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
    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
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
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function







Public Function CustOrderRemarkSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGCustOrderRemark
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from tbCustOrderRemark Order By OrdPos"
    Else
        sqlStr = "Select * from tbCustOrderRemark Where (" & strCond & ") Order By OrdPos"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGCustOrderRemark
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CustOrderRemarkSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

Public Function CR03AddWasteSearch(Optional strCond As String = "") As Collection
On Error GoTo ErrD
Dim tmpItem As BWGJobDataDetail
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim tmpCol As Collection
    If Trim(strCond) = "" Then
        sqlStr = "Select * from vw_CR03AddWasteSearch Order By WorkDate,DocNo"
    Else
        sqlStr = "Select * from vw_CR03AddWasteSearch Where (" & strCond & ") Order By WorkDate,DocNo"
    End If
    
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenKeyset, adLockOptimistic
    Do Until tmpRec.EOF
        Set tmpItem = New BWGJobDataDetail
        tmpItem.ID = Trim("" & tmpRec("JobDetID"))
        tmpItem.MenifestID = Trim("" & tmpRec("MenifestID"))
        tmpItem.MenifestWorkDate = Trim("" & tmpRec("WorkDate"))
        tmpItem.MenifestDocNo = Trim("" & tmpRec("DocNo"))
        tmpItem.WasteName = Trim("" & tmpRec("WasteName"))
        tmpItem.WasteEUCode = Trim("" & tmpRec("EUCode"))
        tmpItem.MenifestCarRegisID = Trim("" & tmpRec("CarRegisID"))
        tmpItem.ChargeWeightType = Trim("" & tmpRec("ChargeWType"))
        tmpItem.DisposerWeight = Trim("" & tmpRec("DisposerWeight"))
        tmpItem.EstWasteQty = Trim("" & tmpRec("EstWasteQty"))
        tmpItem.WasteDataID = Trim("" & tmpRec("WasteDataID"))
        tmpCol.Add tmpItem, "'" & tmpItem.ID & "_" & tmpItem.MenifestID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set CR03AddWasteSearch = tmpCol
    Set tmpCol = Nothing
    Exit Function
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Function
End Function

