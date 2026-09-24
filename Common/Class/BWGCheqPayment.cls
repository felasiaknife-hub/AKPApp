VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGCheqPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim payDate_ As String
Dim cheqNo1_ As String
Dim cheqNo2_ As String
Dim cheqNo3_ As String
Dim cheqNo4_ As String
Dim cheqNo5_ As String
Dim cheqNo6_ As String
Dim cheqNo7_ As String
Dim cheqNo8_ As String
Dim cheqBank1_ As String
Dim cheqBank2_ As String
Dim cheqBank3_ As String
Dim cheqBank4_ As String
Dim cheqBank5_ As String
Dim cheqBank6_ As String
Dim cheqBank7_ As String
Dim cheqBank8_ As String
Dim cheqDate1_ As String
Dim cheqDate2_ As String
Dim cheqDate3_ As String
Dim cheqDate4_ As String
Dim cheqDate5_ As String
Dim cheqDate6_ As String
Dim cheqDate7_ As String
Dim cheqDate8_ As String
Dim cheqAmt1_ As Double
Dim cheqAmt2_ As Double
Dim cheqAmt3_ As Double
Dim cheqAmt4_ As Double
Dim cheqAmt5_ As Double
Dim cheqAmt6_ As Double
Dim cheqAmt7_ As Double
Dim cheqAmt8_ As Double
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get PayDate() As String
    PayDate = payDate_
End Property
Public Property Let PayDate(ByVal newVal As String)
    If Trim(payDate_) <> Trim(newVal) Then
        payDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqNo1() As String
    CheqNo1 = cheqNo1_
End Property
Public Property Let CheqNo1(ByVal newVal As String)
    If Trim(cheqNo1_) <> Trim(newVal) Then
        cheqNo1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqBank1() As String
    CheqBank1 = cheqBank1_
End Property
Public Property Let CheqBank1(ByVal newVal As String)
    If Trim(cheqBank1_) <> Trim(newVal) Then
        cheqBank1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqDate1() As String
    CheqDate1 = cheqDate1_
End Property
Public Property Let CheqDate1(ByVal newVal As String)
    If Trim(cheqDate1_) <> Trim(newVal) Then
        cheqDate1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqAmt1() As Double
    CheqAmt1 = cheqAmt1_
End Property
Public Property Let CheqAmt1(ByVal newVal As Double)
    If cheqAmt1_ <> newVal Then
        cheqAmt1_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqNo2() As String
    CheqNo2 = cheqNo2_
End Property
Public Property Let CheqNo2(ByVal newVal As String)
    If Trim(cheqNo2_) <> Trim(newVal) Then
        cheqNo2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqBank2() As String
    CheqBank2 = cheqBank2_
End Property
Public Property Let CheqBank2(ByVal newVal As String)
    If Trim(cheqBank2_) <> Trim(newVal) Then
        cheqBank2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqDate2() As String
    CheqDate2 = cheqDate2_
End Property
Public Property Let CheqDate2(ByVal newVal As String)
    If Trim(cheqDate2_) <> Trim(newVal) Then
        cheqDate2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqAmt2() As Double
    CheqAmt2 = cheqAmt2_
End Property
Public Property Let CheqAmt2(ByVal newVal As Double)
    If cheqAmt2_ <> newVal Then
        cheqAmt2_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqNo3() As String
    CheqNo3 = cheqNo3_
End Property
Public Property Let CheqNo3(ByVal newVal As String)
    If Trim(cheqNo3_) <> Trim(newVal) Then
        cheqNo3_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqBank3() As String
    CheqBank3 = cheqBank3_
End Property
Public Property Let CheqBank3(ByVal newVal As String)
    If Trim(cheqBank3_) <> Trim(newVal) Then
        cheqBank3_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqDate3() As String
    CheqDate3 = cheqDate3_
End Property
Public Property Let CheqDate3(ByVal newVal As String)
    If Trim(cheqDate3_) <> Trim(newVal) Then
        cheqDate3_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqAmt3() As Double
    CheqAmt3 = cheqAmt3_
End Property
Public Property Let CheqAmt3(ByVal newVal As Double)
    If cheqAmt3_ <> newVal Then
        cheqAmt3_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqNo4() As String
    CheqNo4 = cheqNo4_
End Property
Public Property Let CheqNo4(ByVal newVal As String)
    If Trim(cheqNo4_) <> Trim(newVal) Then
        cheqNo4_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqBank4() As String
    CheqBank4 = cheqBank4_
End Property
Public Property Let CheqBank4(ByVal newVal As String)
    If Trim(cheqBank4_) <> Trim(newVal) Then
        cheqBank4_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqDate4() As String
    CheqDate4 = cheqDate4_
End Property
Public Property Let CheqDate4(ByVal newVal As String)
    If Trim(cheqDate4_) <> Trim(newVal) Then
        cheqDate4_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqAmt4() As Double
    CheqAmt4 = cheqAmt4_
End Property
Public Property Let CheqAmt4(ByVal newVal As Double)
    If cheqAmt4_ <> newVal Then
        cheqAmt4_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqNo5() As String
    CheqNo5 = cheqNo5_
End Property
Public Property Let CheqNo5(ByVal newVal As String)
    If Trim(cheqNo5_) <> Trim(newVal) Then
        cheqNo5_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqBank5() As String
    CheqBank5 = cheqBank5_
End Property
Public Property Let CheqBank5(ByVal newVal As String)
    If Trim(cheqBank5_) <> Trim(newVal) Then
        cheqBank5_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqDate5() As String
    CheqDate5 = cheqDate5_
End Property
Public Property Let CheqDate5(ByVal newVal As String)
    If Trim(cheqDate5_) <> Trim(newVal) Then
        cheqDate5_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqAmt5() As Double
    CheqAmt5 = cheqAmt5_
End Property
Public Property Let CheqAmt5(ByVal newVal As Double)
    If cheqAmt5_ <> newVal Then
        cheqAmt5_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqNo6() As String
    CheqNo6 = cheqNo6_
End Property
Public Property Let CheqNo6(ByVal newVal As String)
    If Trim(cheqNo6_) <> Trim(newVal) Then
        cheqNo6_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqBank6() As String
    CheqBank6 = cheqBank6_
End Property
Public Property Let CheqBank6(ByVal newVal As String)
    If Trim(cheqBank6_) <> Trim(newVal) Then
        cheqBank6_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqDate6() As String
    CheqDate6 = cheqDate6_
End Property
Public Property Let CheqDate6(ByVal newVal As String)
    If Trim(cheqDate6_) <> Trim(newVal) Then
        cheqDate6_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqAmt6() As Double
    CheqAmt6 = cheqAmt6_
End Property
Public Property Let CheqAmt6(ByVal newVal As Double)
    If cheqAmt6_ <> newVal Then
        cheqAmt6_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqNo7() As String
    CheqNo7 = cheqNo7_
End Property
Public Property Let CheqNo7(ByVal newVal As String)
    If Trim(cheqNo7_) <> Trim(newVal) Then
        cheqNo7_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqBank7() As String
    CheqBank7 = cheqBank7_
End Property
Public Property Let CheqBank7(ByVal newVal As String)
    If Trim(cheqBank7_) <> Trim(newVal) Then
        cheqBank7_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqDate7() As String
    CheqDate7 = cheqDate7_
End Property
Public Property Let CheqDate7(ByVal newVal As String)
    If Trim(cheqDate7_) <> Trim(newVal) Then
        cheqDate7_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqAmt7() As Double
    CheqAmt7 = cheqAmt7_
End Property
Public Property Let CheqAmt7(ByVal newVal As Double)
    If cheqAmt7_ <> newVal Then
        cheqAmt7_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get CheqNo8() As String
    CheqNo8 = cheqNo8_
End Property
Public Property Let CheqNo8(ByVal newVal As String)
    If Trim(cheqNo8_) <> Trim(newVal) Then
        cheqNo8_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqBank8() As String
    CheqBank8 = cheqBank8_
End Property
Public Property Let CheqBank8(ByVal newVal As String)
    If Trim(cheqBank8_) <> Trim(newVal) Then
        cheqBank8_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqDate8() As String
    CheqDate8 = cheqDate8_
End Property
Public Property Let CheqDate8(ByVal newVal As String)
    If Trim(cheqDate8_) <> Trim(newVal) Then
        cheqDate8_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get CheqAmt8() As Double
    CheqAmt8 = cheqAmt8_
End Property
Public Property Let CheqAmt8(ByVal newVal As Double)
    If cheqAmt8_ <> newVal Then
        cheqAmt8_ = newVal
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
    sqlStr = "SELECT *, CONVERT(DATETIME, GETDATE(), 103) AS GETDATE FROM tbCheqPayment WHERE PayRefID='" & id_ & "'"
    Set tmpRec = New ADODB.Recordset
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("PayRefID") = id_
    End If
    tmpRec("PayDate") = Trim(payDate_)
    tmpRec("CheqNo1") = Trim(cheqNo1_)
    tmpRec("CheqNo2") = Trim(cheqNo2_)
    tmpRec("CheqNo3") = Trim(cheqNo3_)
    tmpRec("CheqNo4") = Trim(cheqNo4_)
    tmpRec("CheqNo5") = Trim(cheqNo5_)
    tmpRec("CheqNo6") = Trim(cheqNo6_)
    tmpRec("CheqNo7") = Trim(cheqNo7_)
    tmpRec("CheqNo8") = Trim(cheqNo8_)
    tmpRec("CheqBank1") = Trim(cheqBank1_)
    tmpRec("CheqBank2") = Trim(cheqBank2_)
    tmpRec("CheqBank3") = Trim(cheqBank3_)
    tmpRec("CheqBank4") = Trim(cheqBank4_)
    tmpRec("CheqBank5") = Trim(cheqBank5_)
    tmpRec("CheqBank6") = Trim(cheqBank6_)
    tmpRec("CheqBank7") = Trim(cheqBank7_)
    tmpRec("CheqBank8") = Trim(cheqBank8_)
    tmpRec("CheqDate1") = Trim(cheqDate1_)
    tmpRec("CheqDate2") = Trim(cheqDate2_)
    tmpRec("CheqDate3") = Trim(cheqDate3_)
    tmpRec("CheqDate4") = Trim(cheqDate4_)
    tmpRec("CheqDate5") = Trim(cheqDate5_)
    tmpRec("CheqDate6") = Trim(cheqDate6_)
    tmpRec("CheqDate7") = Trim(cheqDate7_)
    tmpRec("CheqDate8") = Trim(cheqDate8_)
    tmpRec("CheqAmt1") = Trim(cheqAmt1_)
    tmpRec("CheqAmt2") = Trim(cheqAmt2_)
    tmpRec("CheqAmt3") = Trim(cheqAmt3_)
    tmpRec("CheqAmt4") = Trim(cheqAmt4_)
    tmpRec("CheqAmt5") = Trim(cheqAmt5_)
    tmpRec("CheqAmt6") = Trim(cheqAmt6_)
    tmpRec("CheqAmt7") = Trim(cheqAmt7_)
    tmpRec("CheqAmt8") = Trim(cheqAmt8_)
    
    tmpRec!LastUser = Trim(Trim(CurrentUser.StaffFName) & " " & Trim(CurrentUser.StaffLName))
    tmpRec!LastUpdate = GetSystemDateTime.SystemDateTime
    
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
    DBConnExc "Delete tbCheqPayment WHERE PayRefID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("PayRefID")
    payDate_ = Trim("" & selRec("PayDate"))
    cheqNo1_ = Trim("" & selRec("CheqNo1"))
    cheqNo2_ = Trim("" & selRec("CheqNo2"))
    cheqNo3_ = Trim("" & selRec("CheqNo3"))
    cheqNo4_ = Trim("" & selRec("CheqNo4"))
    cheqNo5_ = Trim("" & selRec("CheqNo5"))
    cheqNo6_ = Trim("" & selRec("CheqNo6"))
    cheqNo7_ = Trim("" & selRec("CheqNo7"))
    cheqNo8_ = Trim("" & selRec("CheqNo8"))
    cheqBank1_ = Trim("" & selRec("CheqBank1"))
    cheqBank2_ = Trim("" & selRec("CheqBank2"))
    cheqBank3_ = Trim("" & selRec("CheqBank3"))
    cheqBank4_ = Trim("" & selRec("CheqBank4"))
    cheqBank5_ = Trim("" & selRec("CheqBank5"))
    cheqBank6_ = Trim("" & selRec("CheqBank6"))
    cheqBank7_ = Trim("" & selRec("CheqBank7"))
    cheqBank8_ = Trim("" & selRec("CheqBank8"))
    cheqDate1_ = Trim("" & selRec("CheqDate1"))
    cheqDate2_ = Trim("" & selRec("CheqDate2"))
    cheqDate3_ = Trim("" & selRec("CheqDate3"))
    cheqDate4_ = Trim("" & selRec("CheqDate4"))
    cheqDate5_ = Trim("" & selRec("CheqDate5"))
    cheqDate6_ = Trim("" & selRec("CheqDate6"))
    cheqDate7_ = Trim("" & selRec("CheqDate7"))
    cheqDate8_ = Trim("" & selRec("CheqDate8"))
    cheqAmt1_ = Trim("" & selRec("CheqAmt1"))
    cheqAmt2_ = Trim("" & selRec("CheqAmt2"))
    cheqAmt3_ = Trim("" & selRec("CheqAmt3"))
    cheqAmt4_ = Trim("" & selRec("CheqAmt4"))
    cheqAmt5_ = Trim("" & selRec("CheqAmt5"))
    cheqAmt6_ = Trim("" & selRec("CheqAmt6"))
    cheqAmt7_ = Trim("" & selRec("CheqAmt7"))
    cheqAmt8_ = Trim("" & selRec("CheqAmt8"))
End Sub
