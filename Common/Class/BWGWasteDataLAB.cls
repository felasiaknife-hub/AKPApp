VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGWasteDataLAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim wasteDataID_ As String
Dim generateDetail_ As String
Dim wasteChar_ As String
Dim powderOrLump_ As String
Dim CharDescOther_ As String
Dim isPumpAble_ As String
Dim isExplode_ As String
Dim isFlameAble_ As String
Dim flashPoint_ As String
Dim isCorrosive_ As String
Dim phValue_ As String
Dim isReact_ As String
Dim reactWith_ As String
Dim isCarcinogenic_ As String
Dim carcinogenicDesc_ As String
Dim isToxic_ As String
Dim toxicType_ As String
Dim toxicDesc_ As String
Dim isRadioActive_ As String
Dim radioActiveDesc_ As String
Dim isPesticide_ As String
Dim pesticideDesc_ As String
Dim otherDesc_ As String
Dim mixOrSeperate_ As String
Dim inOrOutDoor_ As String
Dim storageType_ As String
Dim storageDesc_ As String
Dim oldEliminator_ As String
Dim colOldTreatment_ As Collection
Dim colNewTreatment_ As Collection

Dim isOdour_ As String
Dim Odour_ As String
Dim isColor_ As String
Dim Color_ As String
Dim isHospital_ As String
Dim isPlasticBag_ As String
Dim PlasticBag_ As String
Dim isPlasticDrum_ As String
Dim PlasticDrum_ As String
Dim isPlasticGallon_ As String
Dim PlasticGallon_ As String
Dim isPaperBox_ As String
Dim PaperBox_ As String
Dim isBottle_ As String
Dim Bottle_ As String

Dim AnalysisType_ As Integer
Dim AnalysisDesc_ As String
Dim isBlending_ As String

Dim isChanged_ As Boolean
Dim isNew_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get wastedataID() As String
    wastedataID = wasteDataID_
End Property
Public Property Let wastedataID(ByVal newVal As String)
    If Trim(wasteDataID_) <> Trim(newVal) Then
        wasteDataID_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get GenerateDetail() As String
    GenerateDetail = generateDetail_
End Property
Public Property Let GenerateDetail(ByVal newVal As String)
    If Trim(generateDetail_) <> Trim(newVal) Then
        generateDetail_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get WasteCharacter() As String
    WasteCharacter = wasteChar_
End Property
Public Property Let WasteCharacter(ByVal newVal As String)
    If Trim(wasteChar_) <> Trim(newVal) Then
        wasteChar_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get PowderOrLump() As String
    PowderOrLump = powderOrLump_
End Property
Public Property Let PowderOrLump(ByVal newVal As String)
    If Trim(powderOrLump_) <> Trim(newVal) Then
        powderOrLump_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CharDescOther() As String
    CharDescOther = CharDescOther_
End Property
Public Property Let CharDescOther(ByVal newVal As String)
    If Trim(CharDescOther_) <> Trim(newVal) Then
        CharDescOther_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isPumpAble() As Boolean
    If isPumpAble_ = "Y" Then
        isPumpAble = True
    Else
        isPumpAble = False
    End If
End Property
Public Property Let isPumpAble(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isPumpAble_) <> Trim(tmpStr) Then
        isPumpAble_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isExplode() As Boolean
    If isExplode_ = "Y" Then
        isExplode = True
    Else
        isExplode = False
    End If
End Property
Public Property Let isExplode(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isExplode_) <> Trim(tmpStr) Then
        isExplode_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get isFlameAble() As Boolean
    If isFlameAble_ = "Y" Then
        isFlameAble = True
    Else
        isFlameAble = False
    End If
End Property
Public Property Let isFlameAble(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isFlameAble_) <> Trim(tmpStr) Then
        isFlameAble_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get FlashPoint() As String
    FlashPoint = flashPoint_
End Property
Public Property Let FlashPoint(ByVal newVal As String)
    If flashPoint_ <> newVal Then
        flashPoint_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get isCorrosive() As Boolean
    If isCorrosive_ = "Y" Then
        isCorrosive = True
    Else
        isCorrosive = False
    End If
End Property
Public Property Let isCorrosive(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isCorrosive_) <> Trim(tmpStr) Then
        isCorrosive_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get pHValue() As String
    pHValue = phValue_
End Property
Public Property Let pHValue(ByVal newVal As String)
    If phValue_ <> newVal Then
        phValue_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get isReact() As Boolean
    If isReact_ = "Y" Then
        isReact = True
    Else
        isReact = False
    End If
End Property
Public Property Let isReact(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isReact_) <> Trim(tmpStr) Then
        isReact_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get ReActWith() As String
    ReActWith = reactWith_
End Property
Public Property Let ReActWith(ByVal newVal As String)
    If Trim(reactWith_) <> Trim(newVal) Then
        reactWith_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isCarcinoGenic() As Boolean
    If isCarcinogenic_ = "Y" Then
        isCarcinoGenic = True
    Else
        isCarcinoGenic = False
    End If
End Property
Public Property Let isCarcinoGenic(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isCarcinogenic_) <> Trim(tmpStr) Then
        isCarcinogenic_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get CarcinoGenicDesc() As String
    CarcinoGenicDesc = carcinogenicDesc_
End Property
Public Property Let CarcinoGenicDesc(ByVal newVal As String)
    If Trim(carcinogenicDesc_) <> Trim(newVal) Then
        carcinogenicDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isToxic() As Boolean
    If isToxic_ = "Y" Then
        isToxic = True
    Else
        isToxic = False
    End If
End Property
Public Property Let isToxic(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isToxic_) <> Trim(tmpStr) Then
        isToxic_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get ToxicType() As String
    ToxicType = toxicType_
End Property
Public Property Let ToxicType(ByVal newVal As String)
    If Trim(toxicType_) <> Trim(newVal) Then
        toxicType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get ToxicDesc() As String
    ToxicDesc = toxicDesc_
End Property
Public Property Let ToxicDesc(ByVal newVal As String)
    If Trim(toxicDesc_) <> Trim(newVal) Then
        toxicDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isRadioActive() As Boolean
    If isRadioActive_ = "Y" Then
        isRadioActive = True
    Else
        isRadioActive = False
    End If
End Property
Public Property Let isRadioActive(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isRadioActive_) <> Trim(tmpStr) Then
        isRadioActive_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get RadioActiveDesc() As String
    RadioActiveDesc = radioActiveDesc_
End Property
Public Property Let RadioActiveDesc(ByVal newVal As String)
    If Trim(radioActiveDesc_) <> Trim(newVal) Then
        radioActiveDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get isPesticide() As Boolean
    If isPesticide_ = "Y" Then
        isPesticide = True
    Else
        isPesticide = False
    End If
End Property
Public Property Let isPesticide(ByVal newVal As Boolean)
Dim tmpStr As String
    If newVal Then
        tmpStr = "Y"
    Else
        tmpStr = "N"
    End If
    If Trim(isPesticide_) <> Trim(tmpStr) Then
        isPesticide_ = Trim(tmpStr)
        isChanged_ = True
    End If
End Property

Public Property Get PesticideDesc() As String
    PesticideDesc = pesticideDesc_
End Property
Public Property Let PesticideDesc(ByVal newVal As String)
    If Trim(pesticideDesc_) <> Trim(newVal) Then
        pesticideDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get OtherDesc() As String
    OtherDesc = otherDesc_
End Property
Public Property Let OtherDesc(ByVal newVal As String)
    If Trim(otherDesc_) <> Trim(newVal) Then
        otherDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get MixOrSeperate() As String
    MixOrSeperate = mixOrSeperate_
End Property
Public Property Let MixOrSeperate(ByVal newVal As String)
    If Trim(mixOrSeperate_) <> Trim(newVal) Then
        mixOrSeperate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get InOrOutDoor() As String
    InOrOutDoor = inOrOutDoor_
End Property
Public Property Let InOrOutDoor(ByVal newVal As String)
    If Trim(inOrOutDoor_) <> Trim(newVal) Then
        inOrOutDoor_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get StorageType() As String
    StorageType = storageType_
End Property
Public Property Let StorageType(ByVal newVal As String)
    If Trim(storageType_) <> Trim(newVal) Then
        storageType_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get StorageDesc() As String
    StorageDesc = storageDesc_
End Property
Public Property Let StorageDesc(ByVal newVal As String)
    If Trim(storageDesc_) <> Trim(newVal) Then
        storageDesc_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get OldEliminator() As String
    OldEliminator = oldEliminator_
End Property
Public Property Let OldEliminator(ByVal newVal As String)
    If Trim(oldEliminator_) <> Trim(newVal) Then
        oldEliminator_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get OldTreatments() As Collection
Dim tmpS As New BWGSearchManager
    If isNew_ = False Then
        If colOldTreatment_ Is Nothing Then Set colOldTreatment_ = tmpS.TreatmentDataSearch("TreatmentID in (Select TreatmentID from tbOldTreatment WHERE WasteDataLabID='" & id_ & "')")
    Else
        If colOldTreatment_ Is Nothing Then Set colOldTreatment_ = tmpS.TreatmentDataSearch("TreatmentID in (Select TreatmentID from tbOldTreatment_New WHERE WasteDataLabID='" & id_ & "')")
    End If
    Set OldTreatments = colOldTreatment_
End Property
Public Property Set OldTreatments(ByVal newVal As Collection)
    Set colOldTreatment_ = newVal
    isChanged_ = True
End Property

Public Property Get NewTreatments() As Collection
Dim tmpS As New BWGSearchManager
    If colNewTreatment_ Is Nothing Then Set colNewTreatment_ = tmpS.TreatmentDataSearch("TreatmentID in (Select TreatmentID from tbTreatment WHERE WasteDataLabID='" & id_ & "')")
    Set NewTreatments = colNewTreatment_
End Property
Public Property Set NewTreatments(ByVal newVal As Collection)
    Set colNewTreatment_ = newVal
    isChanged_ = True
End Property

Public Sub RequeryOldTreatment()
    Set colOldTreatment_ = Nothing
End Sub

Public Sub RequeryNewTreatment()
    Set colNewTreatment_ = Nothing
End Sub

'*********** New Variable 13/06/2551 ******************
Public Property Get isOdour() As Boolean
    If isOdour_ = "Y" Then
        isOdour = True
    Else
        isOdour = False
    End If
End Property
Public Property Let isOdour(ByVal newVal As Boolean)
    If newVal Then
        isOdour_ = "Y"
    Else
        isOdour_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get Odour() As String
    Odour = Odour_
End Property
Public Property Let Odour(ByVal newVal As String)
    If Trim(Odour_) <> Trim(newVal) Then
        Odour_ = Trim(newVal)
        isChanged_ = True
    End If
    isChanged_ = True
End Property

Public Property Get isColor() As Boolean
    If isColor_ = "Y" Then
        isColor = True
    Else
        isColor = False
    End If
End Property
Public Property Let isColor(ByVal newVal As Boolean)
    If newVal Then
        isColor_ = "Y"
    Else
        isColor_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get Color() As String
    Color = Color_
End Property
Public Property Let Color(ByVal newVal As String)
    Color_ = Trim(newVal)
    isChanged_ = True
End Property

Public Property Get isHospital() As Boolean
    If isHospital_ = "Y" Then
        isHospital = True
    Else
        isHospital = False
    End If
End Property
Public Property Let isHospital(ByVal newVal As Boolean)
    If newVal Then
        isHospital_ = "Y"
    Else
        isHospital_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get isPlasticBag() As Boolean
    If isPlasticBag_ = "Y" Then
        isPlasticBag = True
    Else
        isPlasticBag = False
    End If
End Property
Public Property Let isPlasticBag(ByVal newVal As Boolean)
    If newVal Then
        isPlasticBag_ = "Y"
    Else
        isPlasticBag_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get PlasticBag() As String
    PlasticBag = PlasticBag_
End Property
Public Property Let PlasticBag(ByVal newVal As String)
    PlasticBag_ = Trim(newVal)
    isChanged_ = True
End Property

Public Property Get isPlasticDrum() As Boolean
    If isPlasticDrum_ = "Y" Then
        isPlasticDrum = True
    Else
        isPlasticDrum = False
    End If
End Property
Public Property Let isPlasticDrum(ByVal newVal As Boolean)
    If newVal Then
        isPlasticDrum_ = "Y"
    Else
        isPlasticDrum_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get PlasticDrum() As String
    PlasticDrum = PlasticDrum_
End Property
Public Property Let PlasticDrum(ByVal newVal As String)
    PlasticDrum_ = Trim(newVal)
    isChanged_ = True
End Property

Public Property Get isPlasticGallon() As Boolean
    If isPlasticGallon_ = "Y" Then
        isPlasticGallon = True
    Else
        isPlasticGallon = False
    End If
End Property
Public Property Let isPlasticGallon(ByVal newVal As Boolean)
    If newVal Then
        isPlasticGallon_ = "Y"
    Else
        isPlasticGallon_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get PlasticGallon() As String
    PlasticGallon = PlasticGallon_
End Property
Public Property Let PlasticGallon(ByVal newVal As String)
    PlasticGallon_ = Trim(newVal)
    isChanged_ = True
End Property

Public Property Get isPaperBox() As Boolean
    If isPaperBox_ = "Y" Then
        isPaperBox = True
    Else
        isPaperBox = False
    End If
End Property
Public Property Let isPaperBox(ByVal newVal As Boolean)
    If newVal Then
        isPaperBox_ = "Y"
    Else
        isPaperBox_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get PaperBox() As String
    PaperBox = PaperBox_
End Property
Public Property Let PaperBox(ByVal newVal As String)
    PaperBox_ = Trim(newVal)
    isChanged_ = True
End Property

Public Property Get isBottle() As Boolean
    If isBottle_ = "Y" Then
        isBottle = True
    Else
        isBottle = False
    End If
End Property
Public Property Let isBottle(ByVal newVal As Boolean)
    If newVal Then
        isBottle_ = "Y"
    Else
        isBottle_ = "N"
    End If
    isChanged_ = True
End Property

Public Property Get Bottle() As String
    Bottle = Bottle_
End Property
Public Property Let Bottle(ByVal newVal As String)
    Bottle_ = Trim(newVal)
    isChanged_ = True
End Property

Public Property Get AnalysisType() As Integer
    AnalysisType = AnalysisType_
End Property
Public Property Let AnalysisType(ByVal newVal As Integer)
    AnalysisType_ = newVal
    isChanged_ = True
End Property

Public Property Get AnalysisDesc() As String
    AnalysisDesc = AnalysisDesc_
End Property
Public Property Let AnalysisDesc(ByVal newVal As String)
    AnalysisDesc_ = Trim(newVal)
    isChanged_ = True
End Property

Public Property Get isBlending() As String
    isBlending = isBlending_
End Property
Public Property Let isBlending(ByVal newVal As String)
    isBlending_ = newVal
    isChanged_ = True
End Property
'*********** New Variable 13/06/2551 ******************

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Property Get isNew() As Boolean
    isNew = isNew_
End Property
Public Property Let isNew(ByVal newVal As Boolean)
    isNew_ = newVal
End Property

Public Sub Save()
On Error GoTo ErrD
Dim tmpRec As ADODB.Recordset
Dim x%
    GetDBConnection
    Set tmpRec = New ADODB.Recordset
    If isNew_ = False Then
        tmpRec.open "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbWasteDataLab WHERE WasteDataLabID='" & id_ & "'", DBConn, adOpenDynamic, adLockOptimistic
    Else
        tmpRec.open "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbWasteDataLab_New WHERE WasteDataLabID='" & id_ & "'", DBConn, adOpenDynamic, adLockOptimistic
    End If
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("WasteDataLabID") = id_
    End If
    tmpRec("WasteDataID") = Trim(wasteDataID_)
    tmpRec("GenerateDetail") = Trim(generateDetail_)
    tmpRec("WasteCharacter") = Trim(wasteChar_)
    tmpRec("PowderOrLump") = Trim(powderOrLump_)
    tmpRec("CharDescOther") = Trim(CharDescOther_)
    tmpRec("isPumpable") = Trim(isPumpAble_)
    tmpRec("isExplode") = Trim(isExplode_)
    tmpRec("isFlamable") = Trim(isFlameAble_)
    tmpRec("FlashPoint") = Trim(flashPoint_)
    tmpRec("isCorrosive") = Trim(isCorrosive_)
    tmpRec("pHValue") = Trim(phValue_)
    tmpRec("isReact") = Trim(isReact_)
    tmpRec("ReactWith") = Trim(reactWith_)
    tmpRec("isCarcinogenic") = Trim(isCarcinogenic_)
    tmpRec("CarcinogenicDesc") = Trim(carcinogenicDesc_)
    tmpRec("isToxic") = Trim(isToxic_)
    tmpRec("ToxicType") = Trim(toxicType_)
    tmpRec("ToxicDesc") = Trim(toxicDesc_)
    tmpRec("isRadioActive") = Trim(isRadioActive_)
    tmpRec("RadioActiveDesc") = Trim(radioActiveDesc_)
    tmpRec("isPesticide") = Trim(isPesticide_)
    tmpRec("PesticideDesc") = Trim(pesticideDesc_)
    tmpRec("OtherDesc") = Trim(otherDesc_)
    tmpRec("MixOrSeperate") = Trim(mixOrSeperate_)
    tmpRec("InOrOutDoor") = Trim(inOrOutDoor_)
    tmpRec("StorageType") = Trim(storageType_)
    tmpRec("StorageDesc") = Trim(storageDesc_)
    tmpRec("OldEliminator") = Trim(oldEliminator_)
    tmpRec!isHospital = Trim(isHospital_)
    tmpRec!isOdour = Trim(isOdour_)
    tmpRec!Odour = Trim(Odour_)
    tmpRec!isColor = Trim(isColor_)
    tmpRec!Color = Trim(Color_)
    tmpRec!isPlasticBag = Trim(isPlasticBag_)
    tmpRec!PlasticBag = Trim(PlasticBag_)
    tmpRec!isPlasticDrum = Trim(isPlasticDrum_)
    tmpRec!PlasticDrum = Trim(PlasticDrum_)
    tmpRec!isPlasticGallon = Trim(isPlasticGallon_)
    tmpRec!PlasticGallon = Trim(PlasticGallon_)
    tmpRec!isPaperBox = Trim(isPaperBox_)
    tmpRec!PaperBox = Trim(PaperBox_)
    tmpRec!isBottle = Trim(isBottle_)
    tmpRec!Bottle = Trim(Bottle_)
    tmpRec!AnalysisType = Trim(AnalysisType_)
    tmpRec!AnalysisDesc = Trim(AnalysisDesc_)
    tmpRec!isBlending = Trim(isBlending_)
    
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    
    tmpRec.Update
'    tmpRec.Close
    Set tmpRec = Nothing
    
    If Not colOldTreatment_ Is Nothing Then
        If isNew_ = False Then
            DBConnExc "Delete tbOldTreatment WHERE WasteDataLabID='" & id_ & "'"
            For x = 1 To colOldTreatment_.Count
                DBConnExc "Insert into tbOldTreatment (WasteDataLabID,TreatmentID) Values ('" & id_ & "','" & colOldTreatment_(x).ID & "')"
            Next
        Else
            DBConnExc "Delete tbOldTreatment_New WHERE WasteDataLabID='" & id_ & "'"
            For x = 1 To colOldTreatment_.Count
                DBConnExc "Insert into tbOldTreatment_New (WasteDataLabID,TreatmentID) Values ('" & id_ & "','" & colOldTreatment_(x).ID & "')"
            Next
        End If
    End If
    If Not colNewTreatment_ Is Nothing Then
        If isNew_ = False Then
            DBConnExc "Delete tbTreatment WHERE WasteDataLabID='" & id_ & "'"
            For x = 1 To colNewTreatment_.Count
                DBConnExc "Insert into tbTreatment (WasteDataLabID,TreatmentID) Values ('" & id_ & "','" & colNewTreatment_(x).ID & "')"
            Next
        Else
            DBConnExc "Delete tbTreatment_New WHERE WasteDataLabID='" & id_ & "'"
            For x = 1 To colNewTreatment_.Count
                DBConnExc "Insert into tbTreatment_New (WasteDataLabID,TreatmentID) Values ('" & id_ & "','" & colNewTreatment_(x).ID & "')"
            Next
        End If
    End If
    
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
    DBConnExc "Delete tbOldTreatment WHERE WasteDataLabID='" & id_ & "'"
    DBConnExc "Delete tbTreatment WHERE WasteDataLabID='" & id_ & "'"
    DBConnExc "Delete tbWasteDataLab WHERE WasteDataLabID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("WasteDataLabID")
    wasteDataID_ = Trim("" & selRec("WasteDataID"))
    generateDetail_ = Trim("" & selRec("GenerateDetail"))
    wasteChar_ = Trim("" & selRec("WasteCharacter"))
    powderOrLump_ = Trim("" & selRec("PowderOrLump"))
    isPumpAble_ = Trim("" & selRec("isPumpable"))
    isExplode_ = Trim("" & selRec("isExplode"))
    isFlameAble_ = Trim("" & selRec("isFlamable"))
    If IsNull(selRec("FlashPoint")) = False Then flashPoint_ = selRec("FlashPoint")
    isCorrosive_ = Trim("" & selRec("isCorrosive"))
    If IsNull(selRec("pHValue")) = False Then phValue_ = selRec("pHValue")
    isReact_ = Trim("" & selRec("isReact"))
    reactWith_ = Trim("" & selRec("ReactWith"))
    isCarcinogenic_ = Trim("" & selRec("isCarcinogenic"))
    carcinogenicDesc_ = Trim("" & selRec("CarcinogenicDesc"))
    isToxic_ = Trim("" & selRec("isToxic"))
    toxicType_ = Trim("" & selRec("ToxicType"))
    toxicDesc_ = Trim("" & selRec("ToxicDesc"))
    isRadioActive_ = Trim("" & selRec("isRadioActive"))
    radioActiveDesc_ = Trim("" & selRec("RadioActiveDesc"))
    isPesticide_ = Trim("" & selRec("isPesticide"))
    pesticideDesc_ = Trim("" & selRec("PesticideDesc"))
    otherDesc_ = Trim("" & selRec("OtherDesc"))
    mixOrSeperate_ = Trim("" & selRec("MixOrSeperate"))
    inOrOutDoor_ = Trim("" & selRec("InOrOutDoor"))
    storageType_ = Trim("" & selRec("StorageType"))
    storageDesc_ = Trim("" & selRec("StorageDesc"))
    oldEliminator_ = Trim("" & selRec("OldEliminator"))
    
    '******* Add for more waste storage method ************
    isHospital_ = IIf(IsNull(selRec!isHospital), "N", Trim(selRec!isHospital))
    isOdour_ = IIf(IsNull(selRec!isOdour), "N", Trim(selRec!isOdour))
    Odour_ = IIf(IsNull(selRec!Odour), "", Trim(selRec!Odour))
    isColor_ = IIf(IsNull(selRec!isColor), "N", Trim(selRec!isColor))
    Color_ = IIf(IsNull(selRec!Color), "", Trim(selRec!Color))
    isPlasticBag_ = IIf(IsNull(selRec!isPlasticBag), "N", Trim(selRec!isPlasticBag))
    PlasticBag_ = IIf(IsNull(selRec!PlasticBag), "", Trim(selRec!PlasticBag))
    isPlasticDrum_ = IIf(IsNull(selRec!isPlasticDrum), "N", Trim(selRec!isPlasticDrum))
    PlasticDrum_ = IIf(IsNull(selRec!PlasticDrum), "", Trim(selRec!PlasticDrum))
    isPlasticGallon_ = IIf(IsNull(selRec!isPlasticGallon), "N", Trim(selRec!isPlasticGallon))
    PlasticGallon_ = IIf(IsNull(selRec!PlasticGallon), "", Trim(selRec!PlasticGallon))
    isPaperBox_ = IIf(IsNull(selRec!isPaperBox), "N", Trim(selRec!isPaperBox))
    PaperBox_ = IIf(IsNull(selRec!PaperBox), "", Trim(selRec!PaperBox))
    isBottle_ = IIf(IsNull(selRec!isBottle), "N", Trim(selRec!isBottle))
    Bottle_ = IIf(IsNull(selRec!Bottle), "", Trim(selRec!Bottle))
    AnalysisType_ = IIf(IsNull(selRec!AnalysisType), 0, IIf(selRec!AnalysisType = "", 0, selRec!AnalysisType))
    AnalysisDesc_ = IIf(IsNull(selRec!AnalysisDesc), "", Trim(selRec!AnalysisDesc))
    isBlending_ = IIf(IsNull(selRec!isBlending), "", Trim(selRec!isBlending))
    CharDescOther_ = IIf(IsNull(selRec!CharDescOther), "", Trim(selRec!CharDescOther))
End Sub

