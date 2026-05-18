VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGWasteCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WasteCaseID_ As String
Dim CaseName_ As String

Public Property Get WasteCaseID() As String
    WasteCaseID = WasteCaseID_
End Property
Public Property Let WasteCaseID(ByVal newval As String)
    WasteCaseID_ = Trim(newval)
End Property

Public Property Get CaseName() As String
    CaseName = CaseName_
End Property
Public Property Let CaseName(ByVal newval As String)
    CaseName_ = Trim(newval)
End Property


Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    WasteCaseID_ = Trim("" & selRec("WasteCaseID"))
    CaseName_ = Trim("" & selRec("CaseName"))
End Sub
