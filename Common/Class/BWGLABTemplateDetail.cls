VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGLABTemplateDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim templateID_ As String
Dim ParameterID_ As String
Dim paramName_ As String
Dim methodID_ As String
Dim methodName_ As String
Dim labUnitID_ As String
Dim labUnitName_ As String
Dim hasStandard_ As String
Dim standardValue_ As Double
Dim compareType_ As String
Dim commentDesc_ As String
Dim resultType_ As String
Dim itemNo_ As Integer
Dim colLABParam_ As Collection
Dim colLABMethod_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get TemplateID() As String
    TemplateID = templateID_
End Property
Public Property Let TemplateID(ByVal newVal As String)
    If Trim(templateID_) <> Trim(newVal) Then
        templateID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ParameterID() As String
    ParameterID = ParameterID_
End Property
Public Property Let ParameterID(ByVal newVal As String)
    If Trim(ParameterID_) <> Trim(newVal) Then
        ParameterID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ParameterName() As String
    ParameterName = paramName_
End Property
Public Property Let ParameterName(ByVal newVal As String)
    If Trim(paramName_) <> Trim(newVal) Then
        paramName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MethodID() As String
    MethodID = methodID_
End Property
Public Property Let MethodID(ByVal newVal As String)
    If Trim(methodID_) <> Trim(newVal) Then
        methodID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MethodName() As String
    MethodName = methodName_
End Property
Public Property Let MethodName(ByVal newVal As String)
    If Trim(methodName_) <> Trim(newVal) Then
        methodName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LabUnitID() As String
    LabUnitID = labUnitID_
End Property
Public Property Let LabUnitID(ByVal newVal As String)
    If Trim(labUnitID_) <> Trim(newVal) Then
        labUnitID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LabUnitName() As String
    LabUnitName = labUnitName_
End Property
Public Property Let LabUnitName(ByVal newVal As String)
    If Trim(labUnitName_) <> Trim(newVal) Then
        labUnitName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get hasStandard() As Boolean
    If hasStandard_ = "Y" Then
        hasStandard = True
    Else
        hasStandard = False
    End If
End Property
Public Property Let hasStandard(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(hasStandard_) <> Trim(tmpStr) Then
        hasStandard_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get StandardValue() As Double
    StandardValue = standardValue_
End Property

Public Property Let StandardValue(ByVal newVal As Double)
    If standardValue_ <> newVal Then
        standardValue_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CompareType() As String
    CompareType = compareType_
End Property
Public Property Let CompareType(ByVal newVal As String)
    If Trim(compareType_) <> Trim(newVal) Then
        compareType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CommentDesc() As String
    CommentDesc = commentDesc_
End Property
Public Property Let CommentDesc(ByVal newVal As String)
    If Trim(commentDesc_) <> Trim(newVal) Then
        commentDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ResultType() As String
    ResultType = resultType_
End Property
Public Property Let ResultType(ByVal newVal As String)
    If Trim(resultType_) <> Trim(newVal) Then
        resultType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ItemNo() As Integer
    ItemNo = itemNo_
End Property
Public Property Let ItemNo(ByVal newVal As Integer)
    If itemNo_ <> newVal Then
        itemNo_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get LABParam() As Collection
Dim tmpS As New BWGSearchManager
    If colLABParam_ Is Nothing Then Set colLABParam_ = tmpS.ParameterLABSearch("ParameterID='" & ParameterID_ & "'")
    Set LABParam = colLABParam_
End Property
Public Property Set LABParam(ByVal newVal As Collection)
    Set colLABParam_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryLABParam()
    Set colLABParam_ = Nothing
End Sub

Public Property Get LABMethod() As Collection
Dim tmpS As New BWGSearchManager
    If colLABMethod_ Is Nothing Then Set colLABMethod_ = tmpS.MethodLABSearch("MethodID='" & methodID_ & "'")
    Set LABMethod = colLABMethod_
End Property
Public Property Set LABMethod(ByVal newVal As Collection)
    Set colLABMethod_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryLABMethod()
    Set colLABMethod_ = Nothing
End Sub

Public Property Get LABUnit() As Collection
Dim tmpS As New BWGSearchManager
    If colUnit_ Is Nothing Then Set colUnit_ = tmpS.UnitLABSearch("LabUnitID='" & labUnitID_ & "'")
    Set LABUnit = colUnit_
End Property
Public Property Set LABUnit(ByVal newVal As Collection)
    Set colUnit_ = newVal
    isChanged_ = True
End Property
Public Sub RequeryLABUnit()
    Set colUnit_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbLABTemplateDetail WHERE TemplateDetID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("TemplateDetID") = id_
    End If
    tmpRec("TemplateID") = Trim(templateID_)
    tmpRec("ParameterID") = Trim(ParameterID_)
    tmpRec("MethodID") = Trim(methodID_)
    tmpRec("LABUnitID") = Trim(labUnitID_)
    tmpRec("hasStandard") = Trim(hasStandard_)
    tmpRec("StandardValue") = Trim(standardValue_)
    tmpRec("CompareType") = Trim(compareType_)
    tmpRec("CommentDesc") = Trim(commentDesc_)
    tmpRec("ResultType") = Trim(resultType_)
    tmpRec("ItemNo") = itemNo_
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
    DBConnExc "Delete tbLABTemplateDetail WHERE TemplateDetID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
On Error Resume Next
    id_ = selRec("TemplateDetID")
    templateID_ = Trim("" & selRec("TemplateID"))
    ParameterID_ = Trim("" & selRec("ParameterID"))
    paramName_ = Trim("" & selRec("ParameterName"))
    methodID_ = Trim("" & selRec("MethodID"))
    methodName_ = Trim("" & selRec("MethodName"))
    labUnitID_ = Trim("" & selRec("LABUnitID"))
    labUnitName_ = Trim("" & selRec("UnitName"))
    hasStandard_ = Trim("" & selRec("hasStandard"))
    standardValue_ = Trim("" & selRec("StandardValue"))
    compareType_ = Trim("" & selRec("CompareType"))
    commentDesc_ = Trim("" & selRec("CommentDesc"))
    resultType_ = Trim("" & selRec("ResultType"))
    If IsNull(selRec("ItemNo")) = False Then itemNo_ = Trim("" & selRec("ItemNo"))
End Sub



