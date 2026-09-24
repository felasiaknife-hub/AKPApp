VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGApproveLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim approveDate_ As String
Dim approveStaffID_ As String
Dim approveStaffName_ As String
Dim reqStaffID_ As String
Dim reqStaffName_ As String
Dim moduleName_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get ApproveDate() As String
    ApproveDate = Trim(approveDate_)
End Property
Public Property Let ApproveDate(ByVal newVal As String)
    If Trim(approveDate_) <> Trim(newVal) Then
        approveDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ApproveStaffID() As String
    ApproveStaffID = Trim(approveStaffID_)
End Property
Public Property Let ApproveStaffID(ByVal newVal As String)
    If Trim(approveStaffID_) <> Trim(newVal) Then
        approveStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ApproveStaffName() As String
    ApproveStaffName = Trim(approveStaffName_)
End Property
Public Property Let ApproveStaffName(ByVal newVal As String)
    If Trim(approveStaffName_) <> Trim(newVal) Then
        approveStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ReqStaffID() As String
    ReqStaffID = Trim(reqStaffID_)
End Property
Public Property Let ReqStaffID(ByVal newVal As String)
    If Trim(reqStaffID_) <> Trim(newVal) Then
        reqStaffID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ReqStaffName() As String
    ReqStaffName = Trim(reqStaffName_)
End Property
Public Property Let ReqStaffName(ByVal newVal As String)
    If Trim(reqStaffName_) <> Trim(newVal) Then
        reqStaffName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ModuleName() As String
    ModuleName = Trim(moduleName_)
End Property
Public Property Let ModuleName(ByVal newVal As String)
    If Trim(moduleName_) <> Trim(newVal) Then
        moduleName_ = Trim(newVal)
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
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%, sqlStr As String
    sqlStr = "SELECT * FROM tbApproveLog WHERE ApproveLogID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("ApproveLogID") = id_
    End If
    tmpRec("ApproveDate") = Trim(approveDate_)
    tmpRec("ApproveStaffID") = Trim(approveStaffID_)
    tmpRec("ReqStaffID") = Trim(reqStaffID_)
    tmpRec("ModuleName") = Trim(moduleName_)
    tmpRec.Update
'    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
    CloseDBConnection
    Exit Sub
ErrD:
    GetMsg "", vbCritical
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBConnExc "Delete tbApproveLog WHERE ApproveLogID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("ApproveLogID"))
    approveDate_ = Trim("" & selRec("ApproveDate"))
    approveStaffID_ = Trim("" & selRec("ApproveStaffID"))
    approveStaffName_ = Trim("" & selRec("AppStaffFName")) & " " & Trim("" & selRec("AppStaffLName"))
    reqStaffID_ = Trim("" & selRec("ReqStaffID"))
    reqStaffName_ = Trim("" & selRec("ReqStaffFName")) & " " & Trim("" & selRec("ReqStaffLName"))
    moduleName_ = Trim("" & selRec("ModuleName"))
End Sub

