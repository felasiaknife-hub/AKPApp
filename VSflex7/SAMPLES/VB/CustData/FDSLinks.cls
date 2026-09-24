VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlexDataSourceLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' This class will be used as a custom data source
' for the VSFlexGrid control.
'
' It needs to implement the IVSFlexDataSource interface.
'
Implements VSFlex7LCtl.IVSFlexDataSource

Private Function IVSFlexDataSource_GetFieldCount() As Long
    IVSFlexDataSource_GetFieldCount = 9
End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal Field As Long) As String
    Select Case Field
        Case 0: IVSFlexDataSource_GetFieldName = "#"
        Case 1: IVSFlexDataSource_GetFieldName = "ID"
        Case 2: IVSFlexDataSource_GetFieldName = "From"
        Case 3: IVSFlexDataSource_GetFieldName = "To"
        Case 4: IVSFlexDataSource_GetFieldName = "Length"
        Case 5: IVSFlexDataSource_GetFieldName = "Section"
        Case 6: IVSFlexDataSource_GetFieldName = "Nu"
        Case 7: IVSFlexDataSource_GetFieldName = "Nv"
        Case 8: IVSFlexDataSource_GetFieldName = "Nw"
    End Select
End Function

Private Function IVSFlexDataSource_GetRecordCount() As Long
    IVSFlexDataSource_GetRecordCount = UBound(g_Links) - LBound(g_Links) + 1
End Function

Private Function IVSFlexDataSource_GetData(ByVal Field As Long, ByVal Record As Long) As String
    Record = Record + 1 ' we're one-based
    Select Case Field
        Case 0: IVSFlexDataSource_GetData = Record
        Case 1: IVSFlexDataSource_GetData = g_Links(Record).name
        Case 2: IVSFlexDataSource_GetData = g_Links(Record).i
        Case 3: IVSFlexDataSource_GetData = g_Links(Record).j
        Case 5: IVSFlexDataSource_GetData = g_Links(Record).section
        Case 6: IVSFlexDataSource_GetData = g_Links(Record).u
        Case 7: IVSFlexDataSource_GetData = g_Links(Record).v
        Case 8: IVSFlexDataSource_GetData = g_Links(Record).w
        
        ' field 4 is calculated: get the link length
        Case 4:
            Dim dx#, dy#, dz#
            dx = g_Nodes(g_Links(Record).i).x - g_Nodes(g_Links(Record).j).x
            dy = g_Nodes(g_Links(Record).i).y - g_Nodes(g_Links(Record).j).y
            dz = g_Nodes(g_Links(Record).i).z - g_Nodes(g_Links(Record).j).z
            IVSFlexDataSource_GetData = Sqr(dx * dx + dy * dy + dz * dz)
    End Select
End Function

Private Sub IVSFlexDataSource_SetData(ByVal Field As Long, ByVal Record As Long, ByVal newData As String)
    Record = Record + 1 ' we're one-based
    Dim i%
    Select Case Field
        Case 1: ' make sure link name is unique
            For i = LBound(g_Links) To UBound(g_Links)
                If i <> Record And g_Links(Record).name = newData Then
                    MsgBox "Sorry, link names must be unique." & vbCrLf & "This name already belongs to link #" & i & "."
                    Exit Sub
                End If
            Next
            g_Links(Record).name = newData
        
        ' make sure connecting nodes are in the valid range (1 to UBound(g_Nodes), i <> j)
        Case 2:
            i = Val(newData)
            If i < LBound(g_Nodes) Or i > UBound(g_Nodes) Then
                MsgBox "Valid node numbers are between " & LBound(g_Nodes) & " and " & UBound(g_Nodes) & "."
                Exit Sub
            End If
            If i = g_Links(Record).j Then
                MsgBox "Links must connect two different nodes."
                Exit Sub
            End If
            g_Links(Record).i = i
        Case 3:
            i = Val(newData)
            If i < LBound(g_Nodes) Or i > UBound(g_Nodes) Then
                MsgBox "Valid node numbers are between " & LBound(g_Nodes) & " and " & UBound(g_Nodes) & "."
                Exit Sub
            End If
            If i = g_Links(Record).i Then
                MsgBox "Links must connect two different nodes."
                Exit Sub
            End If
            g_Links(Record).j = i
        
        ' section number
        Case 5:
            i = Val(newData)
            If i < 1000 Or i > 2000 Then
                MsgBox "Valid section numbers are between 1000 and 2000."
                Exit Sub
            End If
            g_Links(Record).section = i
            
        ' normal vector needs no validation
        Case 6: g_Links(Record).u = Val(newData)
        Case 7: g_Links(Record).v = Val(newData)
        Case 8: g_Links(Record).w = Val(newData)
        
    End Select

End Sub
