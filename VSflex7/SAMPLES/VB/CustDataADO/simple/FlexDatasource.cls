VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FlexADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' custom FlexGrid data sources implement IVSFlexDataSource
Implements IVSFlexDataSource

' ADO objects and record pointer
Dim m_Conn As New ADODB.Connection
Dim m_RS As New ADODB.Recordset
Dim m_lRecord As Long

' initialize class: create recordset
Private Sub Class_Initialize()
    m_Conn.Open "DSN=NorthWind"
    m_RS.CursorLocation = adUseClient
    m_RS.Open "Orders", m_Conn, adOpenDynamic
    m_lRecord = 1
End Sub

Private Function IVSFlexDataSource_GetData(ByVal Field As Long, ByVal Record As Long) As String

    ' move cursor if we have to
    If m_lRecord <> Record + 1 Then
        m_lRecord = Record + 1
        m_RS.AbsolutePosition = Record + 1
    End If
    
    ' retrieve value
    If Not IsNull(m_RS.Fields(Field).Value) Then
        IVSFlexDataSource_GetData = m_RS.Fields(Field).Value
    End If
    
End Function

Private Function IVSFlexDataSource_GetFieldCount() As Long
    IVSFlexDataSource_GetFieldCount = m_RS.Fields.Count
End Function

Private Function IVSFlexDataSource_GetFieldName(ByVal Field As Long) As String
    IVSFlexDataSource_GetFieldName = m_RS.Fields(Field).Name
End Function

Private Function IVSFlexDataSource_GetRecordCount() As Long
    IVSFlexDataSource_GetRecordCount = m_RS.RecordCount
End Function

Private Sub IVSFlexDataSource_SetData(ByVal Field As Long, ByVal Record As Long, ByVal newData As String)

End Sub
