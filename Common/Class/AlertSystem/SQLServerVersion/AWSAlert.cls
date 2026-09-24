VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AWSAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim alertID_ As String
Dim startDate_ As String
Dim endDate_ As String
Dim manualStop_ As Integer
Dim alertAtStart_ As Integer
Dim alertAtEnd_ As Integer
Dim alertMessage_ As String
Dim isActive_ As Integer
Dim refID_ As String
Dim isChanged_ As Boolean

Private Sub Class_Initialize()
    isChanged_ = False
    isActive_ = 1
End Sub

Public Property Get AlertID() As String
    AlertID = alertID_
End Property
Public Property Let AlertID(ByVal newVal As String)
    alertID_ = Trim(newVal)
End Property

Public Property Get StartDate() As String
    StartDate = startDate_
End Property
Public Property Let StartDate(ByVal newVal As String)
    If Trim(startDate_) <> Trim(newVal) Then
        isChanged_ = True
        startDate_ = Trim(newVal)
    End If
End Property

Public Property Get EndDate() As String
    EndDate = endDate_
End Property
Public Property Let EndDate(ByVal newVal As String)
    If Trim(endDate_) <> Trim(newVal) Then
        isChanged_ = True
        endDate_ = Trim(newVal)
    End If
End Property

Public Property Get RefferenceID() As String
    RefferenceID = refID_
End Property
Public Property Let RefferenceID(ByVal newVal As String)
    If Trim(refID_) <> Trim(newVal) Then
        isChanged_ = True
        refID_ = Trim(newVal)
    End If
End Property

Public Property Get ManualStop() As Integer
    ManualStop = manualStop_
End Property
Public Property Let ManualStop(ByVal newVal As Integer)
    If manualStop_ <> newVal Then
        isChanged_ = True
        manualStop_ = newVal
    End If
End Property

Public Property Get AlertAtStart() As Integer
    AlertAtStart = alertAtStart_
End Property
Public Property Let AlertAtStart(ByVal newVal As Integer)
    If alertAtStart_ <> newVal Then
        isChanged_ = True
        alertAtStart_ = newVal
    End If
End Property

Public Property Get AlertAtEnd() As Integer
    AlertAtEnd = alertAtEnd_
End Property
Public Property Let AlertAtEnd(ByVal newVal As Integer)
    If alertAtEnd_ <> newVal Then
        isChanged_ = True
        alertAtEnd_ = newVal
    End If
End Property

Public Property Get ActiveStatus() As String
    If isActive = 1 Then
        ActiveStatus = "พร้อมเตือน"
    Else
        ActiveStatus = "หยุดเตือน"
    End If
End Property

Public Property Get Message() As String
    Message = ReturnSgQuote(alertMessage_)
End Property
Public Property Let Message(ByVal newVal As String)
    If Trim(alertMessage_) <> Trim(ClearSgQuote(newVal)) Then
        isChanged_ = True
        alertMessage_ = Trim(ClearSgQuote(newVal))
    End If
End Property

Public Property Get MessageSingleLine() As String
    MessageSingleLine = Replace(Replace(ReturnSgQuote(alertMessage_), vbLf, " "), vbCr, " ")
End Property

Public Property Get isActive() As Integer
    isActive = isActive_
End Property
Public Property Let isActive(ByVal newVal As Integer)
    If isActive_ <> newVal Then
        isChanged_ = True
        isActive_ = newVal
    End If
End Property

Public Property Get isChanged() As Boolean
    isChanged = isChanged_
End Property
Public Property Let isChanged(ByVal newVal As Boolean)
    isChanged_ = newVal
End Property

Public Function Save()
On Error GoTo SaveErr:
Dim tmpRec As ADODB.Recordset

    Set tmpRec = New ADODB.Recordset
    If Trim(refID_) = "" Then
'        Set tmpRec = GetRS("Select * from tbAlert Where AlertID='" & alertID_ & "'", adOpenForwardOnly, adLockReadOnly)
        tmpRec.Open "Select * from tbAlert Where AlertID='" & alertID_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    Else
        'tmpRec.Open "Select * from tbAlert Where AlertID='" & alertID_ & "' and ReferenceID='" & refID_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
'        Set tmpRec = GetRS("Select * from tbAlert Where ReferenceID='" & refID_ & "'", adOpenForwardOnly, adLockReadOnly)
        tmpRec.Open "Select * from tbAlert Where ReferenceID='" & refID_ & "'", DBManager.GetDBConnection, adOpenDynamic, adLockOptimistic
    End If
    If tmpRec.EOF Then
        tmpRec.AddNew
        tmpRec("AlertID") = alertID_
    End If
    tmpRec("StartDate") = startDate_
    tmpRec("EndDate") = endDate_
    tmpRec("ManualStop") = manualStop_
    tmpRec("AlertAtStart") = alertAtStart_
    tmpRec("AlertAtEnd") = alertAtEnd_
    tmpRec("AlertMessage") = alertMessage_
    tmpRec("isActive") = isActive_
    tmpRec("ReferenceID") = refID_
    tmpRec.Update
    tmpRec.Close
    Set tmpRec = Nothing
    isChanged_ = False
    Exit Function
    
SaveErr:
    Save = Err

End Function

Public Function Delete() As String
On Error GoTo delErr:
    DBManager.GetDBConnection.Execute "Delete tbAlert Where AlertID='" & alertID_ & "'"
    Delete = ""
    Exit Function
delErr:
    Delete = Err.Description
End Function

Public Sub StopAlert()
    isActive_ = 0
    DBManager.GetDBConnection.Execute "Update tbAlert Set isActive=0 Where AlertID='" & alertID_ & "'"
End Sub

Public Sub SetAlertData(selRec As ADODB.Recordset)
    alertID_ = selRec("AlertID")
    startDate_ = Trim("" & selRec("StartDate"))
    endDate_ = Trim("" & selRec("EndDate"))
    manualStop_ = selRec("ManualStop")
    alertAtStart_ = selRec("AlertAtStart")
    alertAtEnd_ = selRec("AlertAtEnd")
    alertMessage_ = Trim(selRec("AlertMessage"))
    isActive_ = selRec("isActive")
End Sub
