VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFlexGridEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum FlexEditKeyBehaviour
   fgEditNone = 0
   fgEditMoveRight = 1
   fgEditMoveDown = 2
   fgEditMoveup = 3
End Enum

Private Type cntlInfo
      blnIsArray As Boolean
      lngIndex As Long
      strName As String
      blntabStop As Boolean
End Type

Private WithEvents fg As MSFlexGrid
Attribute fg.VB_VarHelpID = -1
Private WithEvents Tb As TextBox
Attribute Tb.VB_VarHelpID = -1

Private strTbName As String
Private prntCntls() As cntlInfo
Private m_TbBorderStyle As AppearanceSettings
Private m_EnterKeyBehaviour As FlexEditKeyBehaviour
Private m_TabKeyBehaviour As FlexEditKeyBehaviour
Private m_blnMoving As Boolean

Public Property Set FlexGridControl(fgControl As MSFlexGrid)
   RemoveOldTextBox
   Set fg = fgControl
   On Error Resume Next
      strTbName = "tbFgEdit"
'      Do
         Err = 0
         strTbName = strTbName & "1"
         fg.Parent.Controls.Add "VB.TextBox", strTbName, fg.Container
'      Loop While Err <> 0
   Set Tb = fg.Parent.Controls(strTbName)
   With Tb
      .Visible = False
      .BorderStyle = m_TbBorderStyle
      Set .Font = fg.Font
      .TabStop = False
      .ZOrder
   End With
End Property

Public Property Let EditBoxBorderStyle(varBorderStyle As AppearanceSettings)
   m_TbBorderStyle = varBorderStyle
   If Not Tb Is Nothing Then Tb.BorderStyle = m_TbBorderStyle
End Property

Public Property Let EnterKeyBehaviour(varKeyBehaviour As FlexEditKeyBehaviour)
   m_EnterKeyBehaviour = varKeyBehaviour
End Property

Public Property Let TabKeyBehaviour(varKeyBehaviour As FlexEditKeyBehaviour)
   ' need to restore tabstops to parent controls if varKeyBehaviour = 0
   If varKeyBehaviour = fgEditNone Then
      If m_TabKeyBehaviour <> fgEditNone Then RestoreTabStops
   End If
   m_TabKeyBehaviour = varKeyBehaviour
End Property

Private Sub Class_Initialize()
   ReDim prntCntls(0)
End Sub

Private Sub Class_Terminate()
   RemoveOldTextBox
End Sub

'++++++++++++++++++++++++++++++
'flex grid events

Private Sub fg_DblClick()
   fgInitEdit 32
End Sub

Private Sub fg_KeyPress(KeyAscii As Integer)
   fgInitEdit KeyAscii
End Sub

Private Sub fg_GotFocus()
   UpdateFg
End Sub

Private Sub fg_LeaveCell()
   If Not m_blnMoving Then UpdateFg
End Sub

Private Sub fg_Scroll()
  Dim dx As Long, dy As Long
  ' move tb with cell or hide if cell out of view
  If Tb.Visible Then
    With fg
      If .RowIsVisible(.Row) And .ColIsVisible(.Col) Then
        dx = .Left + .Container.ScaleX(.CellLeft, vbTwips, .Container.ScaleMode)
        dy = .Top + .Container.ScaleY(.CellTop, vbTwips, .Container.ScaleMode)
        Tb.Move dx, dy
      Else
        Tb.Move -Tb.Width, -Tb.Height
      End If
    End With
  End If
End Sub

'+++++++++++++++++++++++
' Text box events

Private Sub tb_KeyPress(KeyAscii As Integer)
   '  get rid of beeps.
   Select Case KeyAscii
      Case 9, 13, 27
         KeyAscii = 0
   End Select
End Sub

Private Sub tb_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim i As Long
   'read cellTop & cellLeft properties to force into view
   If (Tb.Left < 0) Or (Tb.Top < 0) Then
       i = fg.CellTop
       i = fg.CellLeft
   End If
   
   Select Case KeyCode

   Case 9  ' tab key
      MoveToCell m_TabKeyBehaviour
   
   Case 13   ' ENTER key
      MoveToCell m_EnterKeyBehaviour

   Case 27   ' ESC
      Tb.Visible = False
      fg.SetFocus
 
   Case 38      ' Up.
      MoveToCell fgEditMoveup

   Case 40      ' Down.
      MoveToCell fgEditMoveDown
   End Select

End Sub

Private Sub Tb_LostFocus()
   RestoreTabStops
End Sub

' +++++++++++++++++++++++
' Utility functions

Private Sub fgInitEdit(KeyAscii As Integer)
   Dim sngL As Single, sngT As Single, sngW As Single, sngH As Single
   Select Case KeyAscii
   ' A space means edit the current text.
   Case 0 To 32
      Tb = fg.Text
      Tb.SelStart = 0
      Tb.SelLength = Len(Tb.Text)
   ' Anything else means replace the current text.
   Case Else
      Tb = Chr(KeyAscii)
      Tb.SelStart = Len(Tb.Text)
   End Select
   
   With fg.Container
      sngL = .ScaleX(fg.CellLeft, vbTwips, .ScaleMode)
      sngT = .ScaleY(fg.CellTop, vbTwips, .ScaleMode)
      sngW = .ScaleX(fg.CellWidth, vbTwips, .ScaleMode)
      sngH = .ScaleY(fg.CellHeight, vbTwips, .ScaleMode)
   End With
   ' Show textbox at the right place.
   
   Tb.Move fg.Left + sngL, fg.Top + sngT, sngW, sngH
   Tb.Visible = True
   Tb.SetFocus
   If m_TabKeyBehaviour <> fgEditNone Then
      RestoreTabStops
      RemoveTabStops
   End If
End Sub

Private Sub MoveToCell(varMoveBehaviour As FlexEditKeyBehaviour)
   m_blnMoving = True
   Select Case varMoveBehaviour
   
   Case fgEditNone
      fg.SetFocus
   
   Case fgEditMoveDown
      fg.Text = Tb.Text
      If fg.Row + 1 < fg.Rows Then
         fg.Row = fg.Row + 1
      ElseIf fg.Col + 1 < fg.Cols Then
         fg.Col = fg.Col + 1
         fg.Row = fg.FixedRows
      Else
         fg.Col = fg.FixedCols
         fg.Row = fg.FixedRows
      End If
      fgInitEdit 0
   
   Case fgEditMoveRight
      fg.Text = Tb.Text
      If fg.Col + 1 < fg.Cols Then
         fg.Col = fg.Col + 1
      ElseIf fg.Row + 1 < fg.Rows Then
         fg.Row = fg.Row + 1
         fg.Col = fg.FixedCols
      Else
         fg.Col = fg.FixedCols
         fg.Row = fg.FixedRows
      End If
      fgInitEdit 0

   Case fgEditMoveup
      fg.Text = Tb.Text
      If fg.Row > fg.FixedRows Then
         fg.Row = fg.Row - 1
      ElseIf fg.Col > fg.FixedCols Then
         fg.Col = fg.Col - 1
         fg.Row = fg.Rows - 1
      Else
         fg.Col = fg.Cols - 1
         fg.Row = fg.Rows - 1
      End If
      fgInitEdit 0

   End Select
   m_blnMoving = False
End Sub

Private Sub RemoveTabStops()
   Dim cntl As Control, i As Long
   On Error Resume Next
   With fg.Parent
      ReDim prntCntls(.Controls.Count)
      For Each cntl In .Controls
         i = i + 1
         If TypeName(.Controls(cntl.Name)) = "Object" Then
            prntCntls(i).blnIsArray = True
            prntCntls(i).lngIndex = cntl.Index
         End If
         prntCntls(i).strName = cntl.Name
         prntCntls(i).blntabStop = cntl.TabStop
         cntl.TabStop = False
      Next
   End With
End Sub

Private Sub RestoreTabStops()
   Dim i As Long
   If fg Is Nothing Then Exit Sub
   On Error Resume Next
   With fg.Parent
      For i = 1 To UBound(prntCntls)
         If prntCntls(i).blnIsArray Then
            .Controls(prntCntls(i).strName)(prntCntls(i).lngIndex).TabStop = prntCntls(i).blntabStop
         Else
            .Controls(prntCntls(i).strName).TabStop = prntCntls(i).blntabStop
         End If
      Next
   End With
   ReDim prntCntls(0)
End Sub

Private Sub RemoveOldTextBox()
   On Error Resume Next
   If Not fg Is Nothing Then
     Set Tb = Nothing
     fg.Parent.Controls.Remove strTbName
     Set fg = Nothing
   End If
End Sub

Private Sub UpdateFg()
   If Tb.Visible = False Then Exit Sub
   fg.Text = Tb.Text
   Tb.Visible = False
End Sub



