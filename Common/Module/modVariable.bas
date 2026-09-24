Attribute VB_Name = "modVariable"
Option Explicit

Public RptTypeOrd05 As Integer

Type LogMsg
    LogID As String
    LogJobNo As String
    LogDate As String
    LogByUser As String
    LogCustomerID As String
    LogCustomerSiteID As String
    LogCustomerName As String
    LogCustSiteName As String
    LogMessage As String
    LogState As Integer
    LogForm As Form
End Type

Public NewLog As LogMsg
