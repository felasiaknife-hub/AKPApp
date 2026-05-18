VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApplicationV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim vApp_ As String
Public Property Get vApp() As String
    vApp = vApp_
End Property
Public Property Let vApp(ByVal newval As String)
    vApp_ = Trim(newval)
End Property

Public Sub vAppRecordset(selRec As ADODB.Recordset)
    vApp_ = Trim("" & selRec("vApp"))
End Sub
