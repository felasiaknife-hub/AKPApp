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
Dim signName1_ As String
Dim signPosition1_ As String
Dim signName2_ As String
Dim signPosition2_ As String
Dim signName3_ As String
Dim signPosition3_ As String
Dim isChanged_ As Boolean

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newval As String)
    id_ = Trim(newval)
End Property

Public Property Get SysDate() As String
Dim tmpRec As New ADODB.Recordset
    Set tmpRec = GetRS("Select * from vw_System", adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open "Select * from vw_System", DBManager.GetDBConnection, adOpenStatic, adLockReadOnly
    If tmpRec.EOF = False Then
        sysDate_ = Trim("" & tmpRec("sysdate"))
    End If
    sysDate_ = (CInt(Left(sysDate_, 4)) + 543) & Mid(sysDate_, 5)
    SysDate = sysDate_
End Property
Public Property Let SysDate(ByVal newval As String)
    If Trim(sysDate_) <> Trim(newval) Then
        sysDate_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get VATRate() As Double
    VATRate = vatRate_
End Property
Public Property Let VATRate(ByVal newval As Double)
    If vatRate_ <> newval Then
        vatRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get VendorDeductRate() As Double
    VendorDeductRate = vendorDeductRate_
End Property
Public Property Let VendorDeductRate(ByVal newval As Double)
    If vendorDeductRate_ <> newval Then
        vendorDeductRate_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get VendorFundAmt() As Double
    VendorFundAmt = vendorFundAmt_
End Property
Public Property Let VendorFundAmt(ByVal newval As Double)
    If vendorFundAmt_ <> newval Then
        vendorFundAmt_ = newval
        isChanged_ = True
    End If
End Property

Public Property Get SignName1() As String
    SignName1 = Trim(signName1_)
End Property
Public Property Let SignName1(ByVal newval As String)
    If signName1_ <> Trim(newval) Then
        signName1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignPosition1() As String
    SignPosition1 = Trim(signPosition1_)
End Property
Public Property Let SignPosition1(ByVal newval As String)
    If signPosition1_ <> Trim(newval) Then
        signPosition1_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignName2() As String
    SignName2 = Trim(signName2_)
End Property
Public Property Let SignName2(ByVal newval As String)
    If signName2_ <> Trim(newval) Then
        signName2_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignPosition2() As String
    SignPosition2 = Trim(signPosition2_)
End Property
Public Property Let SignPosition2(ByVal newval As String)
    If signPosition2_ <> Trim(newval) Then
        signPosition2_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignName3() As String
    SignName3 = Trim(signName3_)
End Property
Public Property Let SignName3(ByVal newval As String)
    If signName3_ <> Trim(newval) Then
        signName3_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get SignPosition3() As String
    SignPosition3 = Trim(signPosition3_)
End Property
Public Property Let SignPosition3(ByVal newval As String)
    If signPosition3_ <> Trim(newval) Then
        signPosition3_ = Trim(newval)
        isChanged_ = True
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newval As Boolean)
    isChanged_ = newval
End Property

Public Sub Save()
Dim tmpRec As ADODB.Recordset
    Set tmpRec = New ADODB.Recordset
    If Trim(id_) = "" Then id_ = GetGUID
    tmpRec.Open "Select * from tbSystemConfig Where ConfigID='" & id_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("ConfigID") = id_
    End If
    tmpRec("SysDate") = Trim(sysDate_)
    tmpRec("VATRate") = vatRate_
    tmpRec("VenderDeductRate") = vendorDeductRate_
    tmpRec("VendorFundAmt") = vendorFundAmt_
    tmpRec("SignName1") = signName1_
    tmpRec("SignName2") = signName2_
    tmpRec("SignName3") = signName3_
    tmpRec("SignPosition1") = signPosition1_
    tmpRec("SignPosition2") = signPosition2_
    tmpRec("SignPosition3") = signPosition3_
    tmpRec.Update
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
End Sub

Public Function Delete() As String
On Error GoTo delErr:
    DBManager.GetDBConnection.Execute "Delete tbSystemConfig Where ConfigID='" & id_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub GetSystemConfigFromDB()
Dim tmpRec As New ADODB.Recordset
    Set tmpRec = GetRS("Select * from tbSystemConfig", adOpenForwardOnly, adLockReadOnly)
'    tmpRec.Open "Select * from tbSystemConfig", DBManager.GetDBConnection, adOpenStatic, adLockReadOnly
    If tmpRec.EOF = False Then
        id_ = tmpRec("ConfigID")
        sysDate_ = Trim("" & tmpRec("SysDate"))
        If IsNull(tmpRec("VATRate")) = False Then vatRate_ = tmpRec("VATRate")
        If IsNull(tmpRec("VenderDeductRate")) = False Then vendorDeductRate_ = tmpRec("VenderDeductRate")
        If IsNull(tmpRec("VendorFundAmt")) = False Then vendorFundAmt_ = tmpRec("VendorFundAmt")
        signName1_ = Trim("" & tmpRec("SignName1"))
        signName2_ = Trim("" & tmpRec("SignName2"))
        signName3_ = Trim("" & tmpRec("SignName3"))
        signPosition1_ = Trim("" & tmpRec("SignPosition1"))
        signPosition2_ = Trim("" & tmpRec("SignPosition2"))
        signPosition3_ = Trim("" & tmpRec("SignPosition3"))
    End If
    tmpRec.Close
    Set tmpRec = Nothing
End Sub
