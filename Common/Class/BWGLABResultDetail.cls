VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGLABResultDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim labResultID_ As String
Dim ParameterID_ As String
Dim ParameterName_ As String
Dim methodID_ As String
Dim methodName_ As String
Dim labUnitID_ As String
Dim labUnitName_ As String
Dim hasStandard_ As String
Dim standardValue_ As Double
Dim compareType_ As String
Dim commentDesc_ As String
Dim resultType_ As String
Dim labResultNumber_ As Double
Dim labResultString_ As String
Dim itemNo_ As Integer
Dim testTypeID_ As String
Dim testTypeName_ As String
Dim labResultNumber2_ As Double
Dim labResultString2_ As String
Dim compareType2_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get LABResultID() As String
    LABResultID = labResultID_
End Property
Public Property Let LABResultID(ByVal newVal As String)
    If Trim(labResultID_) <> Trim(newVal) Then
        labResultID_ = Trim(newVal)
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
    ParameterName = ParameterName_
End Property
Public Property Let ParameterName(ByVal newVal As String)
    If Trim(ParameterName_) <> Trim(newVal) Then
        ParameterName_ = Trim(newVal)
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
    If Trim(hasStandard_) = "" Then hasStandard_ = "Y"
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

Public Property Get LABResultNumber() As Double
    LABResultNumber = labResultNumber_
End Property
Public Property Let LABResultNumber(ByVal newVal As Double)
    If labResultNumber_ <> newVal Then
        labResultNumber_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get LABResultString() As String
    LABResultString = labResultString_
End Property
Public Property Let LABResultString(ByVal newVal As String)
    If Trim(labResultString_) <> Trim(newVal) Then
        labResultString_ = Trim(newVal)
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

Public Property Get TestTypeID() As String
    TestTypeID = testTypeID_
End Property
Public Property Let TestTypeID(ByVal newVal As String)
    If Trim(testTypeID_) <> Trim(newVal) Then
        testTypeID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get TestTypeName() As String
    TestTypeName = testTypeName_
End Property
Public Property Let TestTypeName(ByVal newVal As String)
    If Trim(testTypeName_) <> Trim(newVal) Then
        testTypeName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get LABResultNumber2() As Double
    LABResultNumber2 = labResultNumber2_
End Property
Public Property Let LABResultNumber2(ByVal newVal As Double)
    If labResultNumber2_ <> newVal Then
        labResultNumber2_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get LABResultString2() As String
    LABResultString2 = labResultString2_
End Property
Public Property Let LABResultString2(ByVal newVal As String)
    If Trim(labResultString2_) <> Trim(newVal) Then
        labResultString2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CompareType2() As String
    CompareType2 = compareType2_
End Property
Public Property Let CompareType2(ByVal newVal As String)
    If Trim(compareType2_) <> Trim(newVal) Then
        compareType2_ = Trim(newVal)
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
Dim sqlStr As String
    sqlStr = "SELECT * FROM tbLABResultDetail WHERE LABResultDetID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("LABResultDetID") = id_
    End If
    tmpRec("LABResultID") = Trim(labResultID_)
    tmpRec("ParameterID") = Trim(ParameterID_)
    tmpRec("MethodID") = Trim(methodID_)
    tmpRec("LABUnitID") = Trim(labUnitID_)
    tmpRec("hasStandard") = Trim(hasStandard_)
    tmpRec("StandardValue") = Trim(standardValue_)
    tmpRec("CompareType") = Trim(compareType_)
    tmpRec("CommentDesc") = Trim(commentDesc_)
    tmpRec("ResultType") = Trim(resultType_)
    tmpRec("LABResultNumber") = Trim(labResultNumber_)
    tmpRec("LABResultString") = Trim(labResultString_)
    tmpRec("LABResultNumber2") = Trim(labResultNumber2_)
    tmpRec("LABResultString2") = Trim(labResultString2_)
    tmpRec("CompareType2") = Trim(compareType2_)
    tmpRec("TestTypeID") = Trim(testTypeID_)
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
    DBConnExc "Delete tbLABResultDetail WHERE LABResultDetID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
On Error Resume Next
    id_ = selRec("LABResultDetID")
    labResultID_ = Trim("" & selRec("LABResultID"))
    ParameterID_ = Trim("" & selRec("ParameterID"))
    ParameterName_ = Trim("" & selRec("ParameterName"))
    methodID_ = Trim("" & selRec("MethodID"))
    methodName_ = Trim("" & selRec("MethodName"))
    labUnitID_ = Trim("" & selRec("LABUnitID"))
    labUnitName_ = Trim("" & selRec("UnitName"))
    hasStandard_ = Trim("" & selRec("hasStandard"))
    standardValue_ = Trim("" & selRec("StandardValue"))
    compareType_ = Trim("" & selRec("CompareType"))
    commentDesc_ = Trim("" & selRec("CommentDesc"))
    resultType_ = Trim("" & selRec("ResultType"))
    labResultNumber_ = Trim("" & selRec("LABResultNumber"))
    labResultString_ = Trim("" & selRec("LABResultString"))
    compareType2_ = Trim("" & selRec("CompareType2"))
    If IsNull(selRec("LABResultNumber2")) = False Then labResultNumber2_ = Trim("" & selRec("LABResultNumber2"))
    If IsNull(selRec("LABResultString2")) = False Then labResultString2_ = Trim("" & selRec("LABResultString2"))
    testTypeID_ = Trim("" & selRec("TestTypeID"))
    testTypeName_ = Trim("" & selRec("TestTypeName"))
    If IsNull(selRec("ItemNo")) = False Then itemNo_ = Trim("" & selRec("ItemNo"))
End Sub
