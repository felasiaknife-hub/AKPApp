VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGPrintCR03Det"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim cr03ID_ As String
Dim MenifestID_ As String
Dim jobDetID_ As String
Dim workDate_ As String
Dim docNo_ As String
Dim wasteCode_ As String
Dim wasteName_ As String
Dim CarRegisID_ As String
Dim netWeight_ As Double
Dim treatmentDesc_ As String
Dim crRemDesc_ As String
Dim rowPos_ As Integer
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get CR03ID() As String
    CR03ID = cr03ID_
End Property
Public Property Let CR03ID(ByVal newVal As String)
    If Trim(cr03ID_) <> Trim(newVal) Then
        cr03ID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get menifestID() As String
    menifestID = MenifestID_
End Property
Public Property Let menifestID(ByVal newVal As String)
    If Trim(MenifestID_) <> Trim(newVal) Then
        MenifestID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get JobDetID() As String
    JobDetID = jobDetID_
End Property
Public Property Let JobDetID(ByVal newVal As String)
    If Trim(jobDetID_) <> Trim(newVal) Then
        jobDetID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WorkDate() As String
    WorkDate = workDate_
End Property
Public Property Let WorkDate(ByVal newVal As String)
    If Trim(workDate_) <> Trim(newVal) Then
        workDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get docNo() As String
    docNo = docNo_
End Property
Public Property Let docNo(ByVal newVal As String)
    If Trim(docNo_) <> Trim(newVal) Then
        docNo_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteCode() As String
    WasteCode = wasteCode_
End Property
Public Property Let WasteCode(ByVal newVal As String)
    If Trim(wasteCode_) <> Trim(newVal) Then
        wasteCode_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteName() As String
    WasteName = wasteName_
End Property
Public Property Let WasteName(ByVal newVal As String)
    If Trim(wasteName_) <> Trim(newVal) Then
        wasteName_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CarRegisID() As String
    CarRegisID = CarRegisID_
End Property
Public Property Let CarRegisID(ByVal newVal As String)
    If Trim(CarRegisID_) <> Trim(newVal) Then
        CarRegisID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get NetWeight() As Double
    NetWeight = netWeight_
End Property
Public Property Let NetWeight(ByVal newVal As Double)
    If netWeight_ <> newVal Then
        netWeight_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get treatmentDesc() As String
    treatmentDesc = treatmentDesc_
End Property
Public Property Let treatmentDesc(ByVal newVal As String)
    If Trim(treatmentDesc_) <> Trim(newVal) Then
        treatmentDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CRRemDesc() As String
    CRRemDesc = crRemDesc_
End Property
Public Property Let CRRemDesc(ByVal newVal As String)
    If Trim(crRemDesc_) <> Trim(newVal) Then
        crRemDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get RowPos() As Integer
    RowPos = rowPos_
End Property
Public Property Let RowPos(ByVal newVal As Integer)
    If rowPos_ <> newVal Then
        rowPos_ = newVal
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
    sqlStr = "SELECT * FROM tbPrintCR03Detail WHERE CR03DetID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("CR03DetID") = id_
    End If
    tmpRec("CR03ID") = Trim(cr03ID_)
    tmpRec("MenifestID") = Trim(MenifestID_)
    tmpRec("JobDetID") = Trim(jobDetID_)
    tmpRec("CRRemDesc") = Trim(crRemDesc_)
    tmpRec("RowPos") = Trim(rowPos_)
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
    DBConnExc "Delete tbPrintCR03Detail WHERE CR03DetID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("CR03DetID"))
    cr03ID_ = Trim("" & selRec("CR03ID"))
    MenifestID_ = Trim("" & selRec("MenifestID"))
    jobDetID_ = Trim("" & selRec("JobDetID"))
    workDate_ = Trim("" & selRec("WorkDate"))
    docNo_ = Trim("" & selRec("DocNo"))
    wasteCode_ = Trim("" & selRec("EUCode"))
    wasteName_ = Trim("" & selRec("WasteName"))
    CarRegisID_ = Trim("" & selRec("CarRegisID"))
    If Trim("" & selRec("ChargeWType")) = "A" Then 'ผู้รับกำจัด
        If IsNull(selRec("DisposerWeight")) = False Then netWeight_ = selRec("DisposerWeight")
    ElseIf Trim("" & selRec("ChargeWType")) = "B" Then 'ผู้ก่อกำเนิด
        If IsNull(selRec("CustWeight")) = False Then netWeight_ = selRec("CustWeight")
    End If
    treatmentDesc_ = Trim("" & selRec("TreatmentDesc"))
    crRemDesc_ = Trim("" & selRec("CRRemDesc"))
    rowPos_ = selRec("RowPos")
End Sub
