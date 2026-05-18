VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGWasteType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim wasteType_ As String
Dim priorityLevel_ As Integer

Public Property Get WasteType() As String
    WasteType = wasteType_
End Property
Public Property Let WasteType(ByVal newVal As String)
    wasteType_ = Trim(newVal)
End Property

Public Property Get PriorityLevel() As String
    PriorityLevel = priorityLevel_
End Property
Public Property Let PriorityLevel(ByVal newVal As String)
    priorityLevel_ = Trim(newVal)
End Property


Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    wasteType_ = Trim("" & selRec("WasteTypeCode"))
    priorityLevel_ = Trim("" & selRec("PriorityLevel"))
End Sub
