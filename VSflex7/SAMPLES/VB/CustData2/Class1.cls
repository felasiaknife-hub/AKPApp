VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IVSFlexDataSource

Private Function IVSFlexDataSource_GetFieldCount() As Long
    IVSFlexDataSource_GetFieldCount = 4
End Function

Private Function IVSFlexDataSource_GetRecordCount() As Long
    IVSFlexDataSource_GetRecordCount = 1500000
End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal Field As Long) As String
    Select Case Field
        Case 0: IVSFlexDataSource_GetFieldName = "Angle (Deg)"
        Case 1: IVSFlexDataSource_GetFieldName = "Angle (Rad)"
        Case 2: IVSFlexDataSource_GetFieldName = "Sine"
        Case 3: IVSFlexDataSource_GetFieldName = "Co-Sine"
    End Select
End Function

Private Function IVSFlexDataSource_GetData(ByVal Field As Long, ByVal Record As Long) As String
    Select Case Field
        Case 0: IVSFlexDataSource_GetData = Record
        Case 1: IVSFlexDataSource_GetData = Record / 180# * 3.1416
        Case 2: IVSFlexDataSource_GetData = Sin(Record / 180# * 3.1416)
        Case 3: IVSFlexDataSource_GetData = Cos(Record / 180# * 3.1416)
    End Select
End Function

Private Sub IVSFlexDataSource_SetData(ByVal Field As Long, ByVal Record As Long, ByVal newData As String)
    Err.Raise 666, "IVSFlexDataSource", "This data is read-only."
End Sub
