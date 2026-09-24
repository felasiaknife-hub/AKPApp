VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGComRateConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim topHazBonus_ As Double
Dim bestSellerBonus1_ As Double
Dim bestSellerBonus2_ As Double
Dim bestSellerBonus3_ As Double
Dim triMas3PC_ As Integer
Dim triMas3CRBonus_ As Double
Dim triMas3SupBonus_ As Double
Dim newCustNonRate_ As Double
Dim nonMonthPeriod_ As Integer
Dim newCustHzRate_ As Double
Dim hzMonthPeriod_ As Integer
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get TopHazBonus() As Double
    TopHazBonus = topHazBonus_
End Property
Public Property Let TopHazBonus(ByVal newVal As Double)
    If topHazBonus_ <> newVal Then
        topHazBonus_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BestSellerBonus1() As Double
    BestSellerBonus1 = bestSellerBonus1_
End Property
Public Property Let BestSellerBonus1(ByVal newVal As Double)
    If bestSellerBonus1_ <> newVal Then
        bestSellerBonus1_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BestSellerBonus2() As Double
    BestSellerBonus2 = bestSellerBonus2_
End Property
Public Property Let BestSellerBonus2(ByVal newVal As Double)
    If bestSellerBonus2_ <> newVal Then
        bestSellerBonus2_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get BestSellerBonus3() As Double
    BestSellerBonus3 = bestSellerBonus3_
End Property
Public Property Let BestSellerBonus3(ByVal newVal As Double)
    If bestSellerBonus3_ <> newVal Then
        bestSellerBonus3_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get Trimas3PC() As Integer
    Trimas3PC = triMas3PC_
End Property
Public Property Let Trimas3PC(ByVal newVal As Integer)
    If triMas3PC_ <> newVal Then
        triMas3PC_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get Trimas3CRBonus() As Double
    Trimas3CRBonus = triMas3CRBonus_
End Property
Public Property Let Trimas3CRBonus(ByVal newVal As Double)
    If triMas3CRBonus_ <> newVal Then
        triMas3CRBonus_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get Trimas3SupBonus() As Double
    Trimas3SupBonus = triMas3SupBonus_
End Property
Public Property Let Trimas3SupBonus(ByVal newVal As Double)
    If triMas3SupBonus_ <> newVal Then
        triMas3SupBonus_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get NewCustNonRate() As Double
    NewCustNonRate = newCustNonRate_
End Property
Public Property Let NewCustNonRate(ByVal newVal As Double)
    If newCustNonRate_ <> newVal Then
        newCustNonRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get NonMonthPeriod() As Integer
    NonMonthPeriod = nonMonthPeriod_
End Property
Public Property Let NonMonthPeriod(ByVal newVal As Integer)
    If nonMonthPeriod_ <> newVal Then
        nonMonthPeriod_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get NewCustHZRate() As Double
    NewCustHZRate = newCustHzRate_
End Property
Public Property Let NewCustHZRate(ByVal newVal As Double)
    If newCustHzRate_ <> newVal Then
        newCustHzRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get HzMonthPeriod() As Integer
    HzMonthPeriod = hzMonthPeriod_
End Property
Public Property Let HzMonthPeriod(ByVal newVal As Integer)
    If hzMonthPeriod_ <> newVal Then
        hzMonthPeriod_ = newVal
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
    sqlStr = "SELECT * FROM ComRateConfig WHERE RateID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("RateID") = id_
    End If
    tmpRec("TopHazBonus") = topHazBonus_
    tmpRec("BestSalerBonus1") = bestSellerBonus1_
    tmpRec("BestSalerBonus2") = bestSellerBonus2_
    tmpRec("BestSalerBonus3") = bestSellerBonus3_
    tmpRec("TriMas3PC") = triMas3PC_
    tmpRec("TriMas3CRBonus") = triMas3CRBonus_
    tmpRec("TriMas3SupBonus") = triMas3SupBonus_
    tmpRec("NewCustNonRate") = newCustNonRate_
    tmpRec("NonMonthPeriod") = nonMonthPeriod_
    tmpRec("NewCustHZRate") = newCustHzRate_
    tmpRec("HzMonthPeriod") = hzMonthPeriod_
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
    DBConnExc "Delete ComRateConfig WHERE RateID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("RateID")
    If IsNull(selRec("TopHazBonus")) = False Then topHazBonus_ = Trim("" & selRec("TopHazBonus"))
    If IsNull(selRec("BestSalerBonus1")) = False Then bestSellerBonus1_ = Trim("" & selRec("BestSalerBonus1"))
    If IsNull(selRec("BestSalerBonus2")) = False Then bestSellerBonus2_ = Trim("" & selRec("BestSalerBonus2"))
    If IsNull(selRec("BestSalerBonus3")) = False Then bestSellerBonus3_ = Trim("" & selRec("BestSalerBonus3"))
    If IsNull(selRec("TriMas3PC")) = False Then triMas3PC_ = Trim("" & selRec("TriMas3PC"))
    If IsNull(selRec("TriMas3CRBonus")) = False Then triMas3CRBonus_ = Trim("" & selRec("TriMas3CRBonus"))
    If IsNull(selRec("TriMas3SupBonus")) = False Then triMas3SupBonus_ = Trim("" & selRec("TriMas3SupBonus"))
    If IsNull(selRec("NewCustNonRate")) = False Then newCustNonRate_ = Trim("" & selRec("NewCustNonRate"))
    If IsNull(selRec("NonMonthPeriod")) = False Then nonMonthPeriod_ = Trim("" & selRec("NonMonthPeriod"))
    If IsNull(selRec("NewCustHZRate")) = False Then newCustHzRate_ = Trim("" & selRec("NewCustHZRate"))
    If IsNull(selRec("HzMonthPeriod")) = False Then hzMonthPeriod_ = Trim("" & selRec("HzMonthPeriod"))
End Sub
