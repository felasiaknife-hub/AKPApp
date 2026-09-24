VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGBoxDataItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim BoxID_ As String
Dim BoxOwnersCode_ As String
Dim BoxType_ As String
Dim BoxNo_ As String
Dim BoxStatus_ As String
Dim BoxDesc_ As String
Dim isChanged_ As Boolean

Public Property Get BoxID() As String
    BoxID = BoxID_
End Property
Public Property Let BoxID(ByVal newVal As String)
    BoxID_ = Trim(newVal)
End Property

Public Property Get BoxOwnersCode() As String
    BoxOwnersCode = BoxOwnersCode_
End Property
Public Property Let BoxOwnersCode(ByVal newVal As String)
    BoxOwnersCode_ = Trim(newVal)
End Property

Public Property Get BoxType() As String
    BoxType = BoxType_
End Property
Public Property Let BoxType(ByVal newVal As String)
        BoxType_ = Trim(newVal)
End Property

Public Property Get BoxNo() As String
    BoxNo = BoxNo_
End Property
Public Property Let BoxNo(ByVal newVal As String)
        BoxNo_ = Trim(newVal)
End Property

Public Property Get BoxStatus() As String
    BoxStatus = BoxStatus_
End Property
Public Property Let BoxStatus(ByVal newVal As String)
        If BoxStatus_ <> Trim(newVal) Then
            BoxStatus_ = Trim(newVal)
            isChanged_ = True
        End If
End Property

Public Property Get BoxDesc() As String
    BoxDesc = BoxDesc_
End Property
Public Property Let BoxDesc(ByVal newVal As String)
        BoxDesc_ = Trim(newVal)
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
        isChanged_ = newVal
End Property

Public Sub Save()
On Error GoTo ErrHandlers
Dim RS As ADODB.Recordset
Dim tSql As String

    If isChanged_ = False Then Exit Sub
    Set RS = New ADODB.Recordset
    tSql = "SELECT * FROM Box_BoxNo WHERE BoxID = '" & BoxID_ & "'"
    RS.open tSql, DBConStock, adOpenDynamic, adLockOptimistic
    If RS.EOF Then
        RS.AddNew
        RS!BoxID = BoxOwnersCode_ & BoxType_ & BoxNo_
        RS!BoxOwnersCode = BoxOwnersCode_
        RS!BoxType = BoxType_
        RS!BoxCode = BoxNo_
        RS!BoxStatus = BoxStatus_
    End If
    RS!BoxStatus = BoxStatus_
    RS.Update
    
    isChanged_ = False
    Set RS = Nothing
    Exit Sub
    
ErrHandlers:
    MsgBox err.Description
    Exit Sub
End Sub

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    BoxID_ = selRec!BoxID
    BoxOwnersCode_ = selRec!BoxOwnersCode
    BoxType_ = selRec!BoxType
    BoxNo_ = selRec!BoxCode
    BoxStatus_ = selRec!BoxStatus
    BoxDesc_ = Trim("" & selRec!BoxDesc)
End Sub

Public Function BoxDataItemSearch(Optional strCond As String = "", Optional OrderBy As String = "") As Collection
    On Error Resume Next
    Dim tmpItem As BWGBoxDataItem
    Dim tmpRec As ADODB.Recordset
    Dim sqlStr As String
    Dim tmpCol As Collection
    
    sqlStr = "SELECT  *  FROM Box_BoxNo WHERE (" & strCond & ") " & vbCrLf
    If OrderBy = "" Then
        sqlStr = sqlStr & "ORDER BY BoxOwnersCode, BoxType, BoxCode"
    Else
        sqlStr = sqlStr & OrderBy
    End If
    Set tmpRec = New ADODB.Recordset
    Set tmpCol = New Collection
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly, DBConStock)
    Do Until tmpRec.EOF
        Set tmpItem = New BWGBoxDataItem
        tmpItem.SetDataFromRecordset tmpRec
        tmpCol.Add tmpItem, "'" & tmpItem.BoxID & "'"
        Set tmpItem = Nothing
        tmpRec.MoveNext
    Loop
    tmpRec.Close
    Set tmpRec = Nothing
    Set BoxDataItemSearch = tmpCol
    Set tmpCol = Nothing
End Function
