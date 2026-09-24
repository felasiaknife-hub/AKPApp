VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlexDataSourceNodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


' This class will be used as a custom data source
' for the VSFlexGrid control.
'
' It needs to implement the IVSFlexDataSource interface.
'
Implements VSFlex7LCtl.IVSFlexDataSource

Private Function IVSFlexDataSource_GetFieldCount() As Long
    IVSFlexDataSource_GetFieldCount = 5
End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal Field As Long) As String
    Select Case Field
        Case 0: IVSFlexDataSource_GetFieldName = "#"
        Case 1: IVSFlexDataSource_GetFieldName = "ID"
        Case 2: IVSFlexDataSource_GetFieldName = "x"
        Case 3: IVSFlexDataSource_GetFieldName = "y"
        Case 4: IVSFlexDataSource_GetFieldName = "z"
    End Select
End Function

Private Function IVSFlexDataSource_GetRecordCount() As Long
    IVSFlexDataSource_GetRecordCount = UBound(g_Nodes) - LBound(g_Nodes) + 1
End Function

Private Function IVSFlexDataSource_GetData(ByVal Field As Long, ByVal Record As Long) As String
    Record = Record + 1 ' we're one-based
    Select Case Field
        Case 0: IVSFlexDataSource_GetData = Record
        Case 1: IVSFlexDataSource_GetData = g_Nodes(Record).name
        Case 2: IVSFlexDataSource_GetData = g_Nodes(Record).x
        Case 3: IVSFlexDataSource_GetData = g_Nodes(Record).y
        Case 4: IVSFlexDataSource_GetData = g_Nodes(Record).z
    End Select
End Function

Private Sub IVSFlexDataSource_SetData(ByVal Field As Long, ByVal Record As Long, ByVal newData As String)
    Record = Record + 1 ' we're one-based
    Dim i%
    Select Case Field
    
        ' make sure node name is unique
        Case 1:
            For i = LBound(g_Nodes) To UBound(g_Nodes)
                If i <> Record And g_Nodes(i).name = newData Then
                    MsgBox "Sorry, node names must be unique." & vbCrLf & "This name already belongs to node #" & i & "."
                    ' or we could raise an exception:
                    ' err.Raise 666, "Node Data Source", "Non-unique node name"
                    Exit Sub
                End If
            Next
            g_Nodes(Record).name = newData
        
        ' coordinates need no validation
        Case 2: g_Nodes(Record).x = Val(newData)
        Case 3: g_Nodes(Record).y = Val(newData)
        Case 4: g_Nodes(Record).z = Val(newData)
    End Select
End Sub
