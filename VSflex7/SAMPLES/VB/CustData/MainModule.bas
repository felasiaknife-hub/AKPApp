Attribute VB_Name = "MainModule"
Option Explicit

' Here's some custom data structures that will hold the
' data to be displayed and edited with the grid.
'
Type Node_Type
    x As Single     ' node coordinates
    y As Single
    z As Single
    name As String  ' node name
End Type

Type Link_Type
    i As Integer        ' node from
    j As Integer        ' node to
    name As String      ' link name
    section As Integer  ' section type
    u As Single         ' section orientation
    v As Single
    w As Single
End Type


' And this is where the data will be stored.
'
Public g_Nodes(1 To 5000) As Node_Type
Public g_Links(1 To 10000) As Link_Type

' Initialize global data that will be accessed with
' custom objects thet implement IVSFlexDataSource.
'
Sub InitData()

    ' initialize node data
    Dim i%
    For i = LBound(g_Nodes) To UBound(g_Nodes)
        With g_Nodes(i)
            .x = Sin(i / 100) * 500 + Rnd * 200
            .y = Cos(i / 100) * 500 + Rnd * 200
            .z = i / 10
            .name = "Node " & i
        End With
    Next
    
    ' initialize link data
    For i = LBound(g_Nodes) To UBound(g_Links)
        With g_Links(i)
            .i = 1 + (Rnd * 5000) Mod UBound(g_Nodes)
            Do
                .j = 1 + (Rnd * 5000) Mod UBound(g_Nodes)
            Loop Until .j <> .i
            .name = "Link " & i
            .section = 1000 + Int(Rnd * 1000)
            .u = Rnd
            .v = Rnd
            .w = Rnd
        End With
    Next
End Sub
