VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VSDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements VSFlex7LCtl.IVSDataSource

Private Function IVSDataSource_GetData(ByVal Row As Long, ByVal Col As Long) As String
    If Row = 0 And Col = 0 Then
        IVSDataSource_GetData = "*"
    ElseIf Row = 0 Then
        IVSDataSource_GetData = Chr(Asc("A") + Col - 1)
    ElseIf Col = 0 Then
        IVSDataSource_GetData = "R" & Row
    Else
        IVSDataSource_GetData = "R" & Row & " C" & Col
    End If
End Function

Private Sub IVSDataSource_SetData(ByVal Row As Long, ByVal Col As Long, ByVal newVal As String)
    Debug.Print "Setting "; newVal
End Sub
