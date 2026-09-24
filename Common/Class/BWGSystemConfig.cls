VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGSystemConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim sysDate_ As String
Dim vatRate_ As Double
Dim vendorDeductRate_ As Double
Dim vendorFundAmt_ As Double
Dim SignName1_ As String
Dim SignPosition1_ As String
Dim signName2_ As String
Dim signPosition2_ As String
Dim signName3_ As String
Dim signPosition3_ As String
Dim signName4_ As String
Dim signPosition4_ As String
Dim signName5_ As String
Dim signPosition5_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get SysDate() As String
Dim tmpRec As New ADODB.Recordset
Dim sqlStr As String
    sqlStr = "SELECT * FROM vw_System"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open sqlStr, DBConn, adOpenKeyset, adLockOptimistic
    If tmpRec.EOF = False Then
        sysDate_ = Trim("" & tmpRec("sysdate"))
    End If
    sysDate_ = (CInt(Left(sysDate_, 4)) + 543) & Mid(sysDate_, 5)
    SysDate = sysDate_
End Property
Public Property Let SysDate(ByVal newVal As String)
    If Trim(sysDate_) <> Trim(newVal) Then
        sysDate_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get VATRate() As Double
    VATRate = vatRate_
End Property
Public Property Let VATRate(ByVal newVal As Double)
    If vatRate_ <> newVal Then
        vatRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get VendorDeductRate() As Double
    VendorDeductRate = vendorDeductRate_
End Property
Public Property Let VendorDeductRate(ByVal newVal As Double)
    If vendorDeductRate_ <> newVal Then
        vendorDeductRate_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get VendorFundAmt() As Double
    VendorFundAmt = vendorFundAmt_
End Property
Public Property Let VendorFundAmt(ByVal newVal As Double)
    If vendorFundAmt_ <> newVal Then
        vendorFundAmt_ = newVal
        isChanged_ = True
    End If
End Property

Public Property Get SignName1() As String
    SignName1 = Trim(SignName1_)
End Property
Public Property Let SignName1(ByVal newVal As String)
    If SignName1_ <> Trim(newVal) Then
        SignName1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignPosition1() As String
    SignPosition1 = Trim(SignPosition1_)
End Property
Public Property Let SignPosition1(ByVal newVal As String)
    If SignPosition1_ <> Trim(newVal) Then
        SignPosition1_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignName2() As String
    SignName2 = Trim(signName2_)
End Property
Public Property Let SignName2(ByVal newVal As String)
    If signName2_ <> Trim(newVal) Then
        signName2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignPosition2() As String
    SignPosition2 = Trim(signPosition2_)
End Property
Public Property Let SignPosition2(ByVal newVal As String)
    If signPosition2_ <> Trim(newVal) Then
        signPosition2_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignName3() As String
    SignName3 = Trim(signName3_)
End Property
Public Property Let SignName3(ByVal newVal As String)
    If signName3_ <> Trim(newVal) Then
        signName3_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignPosition3() As String
    SignPosition3 = Trim(signPosition3_)
End Property
Public Property Let SignPosition3(ByVal newVal As String)
    If signPosition3_ <> Trim(newVal) Then
        signPosition3_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignName4() As String
    SignName4 = Trim(signName4_)
End Property
Public Property Let SignName4(ByVal newVal As String)
    If signName4_ <> Trim(newVal) Then
        signName4_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignPosition4() As String
    SignPosition4 = Trim(signPosition4_)
End Property
Public Property Let SignPosition4(ByVal newVal As String)
    If signPosition4_ <> Trim(newVal) Then
        signPosition4_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignName5() As String
    SignName5 = Trim(signName5_)
End Property
Public Property Let SignName5(ByVal newVal As String)
    If signName5_ <> Trim(newVal) Then
        signName5_ = Trim(newVal)
        isChanged_ = True
    End If
End Property

Public Property Get SignPosition5() As String
    SignPosition5 = Trim(signPosition5_)
End Property
Public Property Let SignPosition5(ByVal newVal As String)
    If signPosition5_ <> Trim(newVal) Then
        signPosition5_ = Trim(newVal)
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
    Set tmpRec = New ADODB.Recordset
    If Trim(id_) = "" Then id_ = GetGUID
    sqlStr = "SELECT * FROM tbSystemConfig WHERE ConfigID='" & id_ & "'"
    GetDBConnection
    tmpRec.open sqlStr, DBConn, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("ConfigID") = id_
    End If
    tmpRec("SysDate") = Trim(sysDate_)
    tmpRec("VATRate") = vatRate_
    tmpRec("VenderDeductRate") = vendorDeductRate_
    tmpRec("VendorFundAmt") = vendorFundAmt_
    tmpRec("SignName1") = SignName1_
    tmpRec("SignName2") = signName2_
    tmpRec("SignName3") = signName3_
    tmpRec("SignPosition1") = SignPosition1_
    tmpRec("SignPosition2") = signPosition2_
    tmpRec("SignPosition3") = signPosition3_
    tmpRec("SignPosition4") = signPosition4_
    tmpRec("SignPosition5") = signPosition5_
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
    DBConnExc "Delete tbSystemConfig WHERE ConfigID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = err.Description
End Function

Public Sub GetSystemConfigFromDB()
On Error GoTo ErrD
Dim tmpRec As New ADODB.Recordset
Dim sqlStr As String
    GetDBConnection
    sqlStr = "Select top(10) * from tbSystemConfig"
    Set tmpRec = GetRS(sqlStr, adOpenForwardOnly, adLockReadOnly)
    If tmpRec.EOF = False Then
        id_ = tmpRec("ConfigID")
        sysDate_ = Trim("" & tmpRec("SysDate"))
        If IsNull(tmpRec("VATRate")) = False Then vatRate_ = tmpRec("VATRate")
        If IsNull(tmpRec("VenderDeductRate")) = False Then vendorDeductRate_ = tmpRec("VenderDeductRate")
        If IsNull(tmpRec("VendorFundAmt")) = False Then vendorFundAmt_ = tmpRec("VendorFundAmt")
        SignName1_ = Trim("" & tmpRec("SignName1"))
        signName2_ = Trim("" & tmpRec("SignName2"))
        signName3_ = Trim("" & tmpRec("SignName3"))
        signName4_ = Trim("" & tmpRec("SignName4"))
        signName5_ = Trim("" & tmpRec("SignName5"))
        SignPosition1_ = Trim("" & tmpRec("SignPosition1"))
        signPosition2_ = Trim("" & tmpRec("SignPosition2"))
        signPosition3_ = Trim("" & tmpRec("SignPosition3"))
        signPosition4_ = Trim("" & tmpRec("SignPosition4"))
        signPosition5_ = Trim("" & tmpRec("SignPosition5"))
    End If
    tmpRec.Close
    Set tmpRec = Nothing
    Exit Sub
ErrD:
    MsgBox err.Description, vbCritical, "Error"
    End
End Sub

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = selRec("ConfigID")
    SignName1_ = Trim("" & selRec("SignName1"))
'    signName2_ = Trim("" & selRec("SignName2"))
'    signName3_ = Trim("" & selRec("SignName3"))
'    signName4_ = Trim("" & selRec("SignName4"))
'    signName5_ = Trim("" & selRec("SignName5"))
    SignPosition1_ = Trim("" & selRec("SignPosition1"))
'    signPosition2_ = Trim("" & selRec("SignPosition2"))
'    signPosition3_ = Trim("" & selRec("SignPosition3"))
'    signPosition4_ = Trim("" & selRec("SignPosition4"))
'    signPosition5_ = Trim("" & selRec("SignPosition5"))
End Sub
