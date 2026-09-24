VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGTripMinWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim TimeTableNo_ As String
Dim MenifestID_ As String
Dim isUnderWeight_ As Boolean
Dim MaxWeight_ As Double
Dim MaxWasteID_ As String
Dim MaxWasteType_ As String
Dim TotalWeight_ As Double
Dim MaxWeightPerCar_ As Double
Dim MaxRate_ As Double
Dim MaxRateType_ As String
Dim isOverWeight_ As Boolean
Dim colWasteWeight_ As Collection

Public Property Get TimeTableNo() As String
    TimeTableNo = TimeTableNo_
End Property
Public Property Let TimeTableNo(ByVal newVal As String)
    TimeTableNo_ = Trim(newVal)
End Property

Public Property Get menifestID() As String
    menifestID = MenifestID_
End Property
Public Property Let menifestID(ByVal newVal As String)
    MenifestID_ = Trim(newVal)
End Property

Public Property Get isUnderWeight() As Boolean
    isUnderWeight = isUnderWeight_
End Property
Public Property Let isUnderWeight(ByVal newVal As Boolean)
    isUnderWeight_ = Trim(newVal)
End Property

Public Property Get isOverWeight() As Boolean
    isOverWeight = isOverWeight_
End Property
Public Property Let isOverWeight(ByVal newVal As Boolean)
    isOverWeight_ = Trim(newVal)
End Property

Public Property Get MaxWeight() As Double
    MaxWeight = MaxWeight_
End Property
Public Property Let MaxWeight(ByVal newVal As Double)
    MaxWeight_ = Trim(newVal)
End Property

Public Property Get MaxWasteID() As String
    MaxWasteID = MaxWasteID_
End Property
Public Property Let MaxWasteID(ByVal newVal As String)
    MaxWasteID_ = Trim(newVal)
End Property

Public Property Get MaxWasteType() As String
    MaxWasteType = MaxWasteType_
End Property
Public Property Let MaxWasteType(ByVal newVal As String)
    MaxWasteType_ = Trim(newVal)
End Property

Public Property Get TotalWeight() As Double
    TotalWeight = TotalWeight_
End Property
Public Property Let TotalWeight(ByVal newVal As Double)
    TotalWeight_ = Trim(newVal)
End Property

Public Property Get ColWasteWeight() As Collection
    Set ColWasteWeight = colWasteWeight_
End Property
Public Property Set ColWasteWeight(ByVal newVal As Collection)
    Set colWasteWeight_ = newVal
End Property

Public Property Get MaxWeightPerCar() As Double
    MaxWeightPerCar = MaxWeightPerCar_
End Property
Public Property Let MaxWeightPerCar(ByVal newVal As Double)
    MaxWeightPerCar_ = Trim(newVal)
End Property

Public Property Get MaxRate() As Double
    MaxRate = MaxRate_
End Property
Public Property Let MaxRate(ByVal newVal As Double)
    MaxRate_ = Trim(newVal)
End Property

Public Property Get MaxRateType() As String
    MaxRateType = MaxRateType_
End Property
Public Property Let MaxRateType(ByVal newVal As String)
    MaxRateType_ = Trim(newVal)
End Property
