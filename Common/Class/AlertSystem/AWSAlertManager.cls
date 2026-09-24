VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AWSAlertManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim curAlerts_ As Collection

Public Function getActiveAlerts() As Collection
    'LoadAlertFromDB "isActive=1 And (ManualStop=1 Or EndDate>='" & TodayDate & "')"
    LoadAlertFromDB "isActive=1"
    Set getActiveAlerts = curAlerts_
End Function

Public Function getUnActiveAlerts() As Collection
    LoadAlertFromDB "isActive=0"
    Set getUnActiveAlerts = curAlerts_
End Function

Public Function getAllAlerts() As Collection
    LoadAlertFromDB ""
    Set getAllAlerts = curAlerts_
End Function

Public Function getActiveAlertsInDate(startDate1 As String, Optional startDate2 As String = "") As Collection
    If Trim(startDate1) = "" And Trim(startDate2) = "" Then
        getActiveAlerts
    ElseIf Trim(startDate1) <> "" And Trim(startDate2) = "" And IsDate(startDate1) Then
        'LoadAlertFromDB "StartDate='" & startDate1 & "' And isActive=1 And (ManualStop=1 Or EndDate>='" & TodayDate & "')"
        LoadAlertFromDB "StartDate<='" & startDate1 & "' And isActive=1"
    ElseIf Trim(startDate1) <> "" And Trim(startDate2) <> "" And IsDate(startDate1) And IsDate(startDate2) Then
        'LoadAlertFromDB "StartDate>='" & startDate1 & "' And StartDate<='" & startDate2 & "' And isActive=1 And (ManualStop=1 Or EndDate<='" & TodayDate & "')"
        LoadAlertFromDB "StartDate>='" & startDate1 & "' And StartDate<='" & startDate2 & "' And isActive=1"
    End If
    Set getActiveAlertsInDate = curAlerts_
End Function

Private Sub LoadAlertFromDB(strCri As String)
Dim tmpAlert As AWSAlert
Dim tmpRec As ADODB.Recordset
Dim sqlStr As String
Dim x%
    Set curAlerts_ = Nothing
    Set curAlerts_ = New Collection
    Set tmpRec = New ADODB.Recordset
    If Trim(strCri) <> "" Then
        sqlStr = "Select * from tbAlert Where " & strCri & " Order By StartDate"
    Else
        sqlStr = "Select * from tbAlert Order By StartDate"
    End If
    tmpRec.Open sqlStr, DBManager.GetDBConnection, adOpenStatic, adLockReadOnly
    Do Until tmpRec.EOF
        Set tmpAlert = New AWSAlert
        tmpAlert.SetAlertData tmpRec
        curAlerts_.Add tmpAlert, "'" & tmpAlert.AlertID & "'"
        Set tmpAlert = Nothing
        tmpRec.MoveNext
    Loop
End Sub

Private Sub Class_Terminate()
    Set curAlerts_ = Nothing
End Sub

Public Sub ShowAlertAtStart()
Dim x%
    'LoadAlertFromDB "isActive=1 And AlertAtStart=1 And (ManualStop=1 Or EndDate<='" & TodayDate & "')"
    LoadAlertFromDB "isActive=1 And AlertAtStart=1 and StartDate<='" & TodayDate & "'"
    If curAlerts_ Is Nothing Then Exit Sub
    For x = 1 To curAlerts_.Count
        frAlertShow.setAlertItem curAlerts_(x)
        frAlertShow.Show 1
    Next
End Sub

Public Sub ShowAlertAtEnd()
Dim x%
    'LoadAlertFromDB "isActive=1 And AlertAtEnd=1 And (ManualStop=1 Or EndDate<='" & TodayDate & "')"
    LoadAlertFromDB "isActive=1 And AlertAtEnd=1 and StartDate<='" & TodayDate & "'"
    If curAlerts_ Is Nothing Then Exit Sub
    For x = 1 To curAlerts_.Count
        frAlertShow.setAlertItem curAlerts_(x)
        frAlertShow.Show 1
    Next
End Sub

Public Sub AddAutoAlert(alertMsg As String, startDate As String, refID As String, Optional addInd As String)
Dim tmpAlert As AWSAlert
    If Trim(alertMsg) = "" Then Exit Sub
    If Trim(startDate) = "" Then Exit Sub

    Set tmpAlert = New AWSAlert
    With tmpAlert
        .AlertID = getNewProductID(modPrefix.AlertAutoPrefix) & addInd  'Alert Auto
        .AlertAtStart = 1
        .AlertAtEnd = 1
        .ManualStop = 1
        .startDate = startDate
        .Message = alertMsg
        .RefferenceID = refID
        .Save
    End With
    Set tmpAlert = Nothing
End Sub
