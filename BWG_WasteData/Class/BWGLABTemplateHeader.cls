VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGLABTemplateHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim templateNo_ As String
Dim templateName_ As String
Dim createdDate_ As String
Dim testTypeID_ As String
Dim testTypeName_ As String
Dim colLABTempDet_ As Collection
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get TemplateNo() As String
    TemplateNo = templateNo_
End Property
Public Property Let TemplateNo(ByVal newval As String)
    If Trim(templateNo_) <> Trim(newval) Then
        templateNo_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TemplateName() As String
    TemplateName = templateName_
End Property
Public Property Let TemplateName(ByVal newval As String)
    If Trim(templateName_) <> Trim(newval) Then
        templateName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get CreatedDate() As String
    CreatedDate = createdDate_
End Property
Public Property Let CreatedDate(ByVal newval As String)
    If Trim(createdDate_) <> Trim(newval) Then
        createdDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TestTypeID() As String
    TestTypeID = testTypeID_
End Property
Public Property Let TestTypeID(ByVal newval As String)
    If Trim(testTypeID_) <> Trim(newval) Then
        testTypeID_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get TestTypeName() As String
    TestTypeName = testTypeName_
End Property
Public Property Let TestTypeName(ByVal newval As String)
    If Trim(testTypeName_) <> Trim(newval) Then
        testTypeName_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get LABTemplateDet() As Collection
Dim tmpS As New BWGSearchManager
    If colLABTempDet_ Is Nothing Then Set colLABTempDet_ = tmpS.LABTemplateSearch("TemplateID='" & id_ & "'")
    Set LABTemplateDet = colLABTempDet_
End Property
Public Property Set LABTemplateDet(ByVal newval As Collection)
    Set colLABTempDet_ = newval
    isChanged_ = True
End Property
Public Sub RequeryLABTemplateDet()
    Set colLABTempDet_ = Nothing
End Sub

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%

    Set tmpRec = New ADODB.Recordset
    tmpRec.Open "Select * from tbLABTemplateHeader Where TemplateID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("TemplateID") = id_
    End If
    
    tmpRec("TemplateName") = Trim(templateName_)
    tmpRec("TestTypeID") = Trim(testTypeID_)
    tmpRec("CreatedDate") = Trim(createdDate_)
    tmpRec.Update
    tmpRec.Requery
    templateNo_ = Trim("" & tmpRec("TemplateNo"))
    templateNo_ = Trim("" & tmpRec("TemplateNo"))
    tmpRec.Close
        
    With DBManager.GetDBConnection
        .Execute "Delete tbLABTemplateDetail Where TemplateID='" & id_ & "'"
        If Not colLABTempDet_ Is Nothing Then
            For x = 1 To colLABTempDet_.Count
                colLABTempDet_(x).Save
            Next
        End If
    End With

    Set tmpRec = Nothing
    isChanged_ = False
    Exit Sub
ErrD:
    MsgBox Err.Description & vbCrLf & "กรุณาลองใหม่หรือติดต่อผู้พัฒนาระบบ", vbCritical, "Error"
    Exit Sub
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    With DBManager.GetDBConnection
        .Execute "Delete tbLABTemplateDetail Where TemplateID='" & id_ & "'"
        .Execute "Delete tbLABTemplateHeader Where TemplateID='" & id_ & "'"
    End With
    Delete = ""
    Exit Function
    Exit Function
delErr:
    Delete = Err.Description
End Function
Public Sub SetDataFromRecordsetFrom(selRec As ADODB.Recordset)
    id_ = selRec("TemplateID")
    templateNo_ = Trim("" & selRec("TemplateNo"))
    templateName_ = Trim("" & selRec("TemplateName"))
    createdDate_ = Trim("" & selRec("CreatedDate"))
    testTypeID_ = Trim("" & selRec("TestTypeID"))
'    Cost_ = Trim("" & selRec("Cost"))
End Sub

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("TemplateID")
    templateNo_ = Trim("" & selRec("TemplateNo"))
    templateName_ = Trim("" & selRec("TemplateName"))
    createdDate_ = Trim("" & selRec("CreatedDate"))
    testTypeID_ = Trim("" & selRec("TestTypeID"))
    testTypeName_ = Trim("" & selRec("TestTypeName"))
End Sub


