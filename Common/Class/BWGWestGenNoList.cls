VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BWGWasteGenNoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim id_ As String
Dim jobID_ As String
Dim wasteDataID_ As String
Dim wasteCreateNo_ As String

Public Property Get ID() As String
    ID = id_
End Property
Public Property Let ID(ByVal newVal As String)
    id_ = Trim(newVal)
End Property

Public Property Get JobID() As String
    JobID = jobID_
End Property
Public Property Let JobID(ByVal newVal As String)
    jobID_ = Trim(newVal)
End Property

Public Property Get WasteDataID() As String
    WasteDataID = wasteDataID_
End Property
Public Property Let WasteDataID(ByVal newVal As String)
    wasteDataID_ = Trim(newVal)
End Property

Public Property Get WasteCreateNo() As String
    WasteCreateNo = wasteCreateNo_
End Property
Public Property Let WasteCreateNo(ByVal newVal As String)
    wasteCreateNo_ = Trim(newVal)
End Property

Public Sub SetDataFromRecordset(selRec As ADODB.Recordset)
    id_ = Trim("" & selRec("JobDetID"))
    jobID_ = Trim("" & selRec("JobID"))
    wasteDataID_ = Trim("" & selRec("WasteDataID"))
    wasteCreateNo_ = Trim("" & selRec("WasteCreateNo"))
End Sub
