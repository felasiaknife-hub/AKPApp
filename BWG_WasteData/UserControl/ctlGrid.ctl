VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl ctlGrid 
   BackColor       =   &H8000000C&
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10185
   KeyPreview      =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   10185
   Begin MSComCtl2.DTPicker dtInput 
      Height          =   315
      Left            =   1410
      TabIndex        =   3
      Tag             =   "11"
      Top             =   1500
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3735555
      CurrentDate     =   37339
   End
   Begin VB.ComboBox cboInput 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   4500
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2340
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtInput 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   4500
      TabIndex        =   1
      Top             =   750
      Visible         =   0   'False
      Width           =   1245
   End
   Begin MSFlexGridLib.MSFlexGrid fgGrid 
      Height          =   4905
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   8652
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      RowHeightMin    =   315
      BackColorBkg    =   12632256
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "ctlGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Event BeforeUpdateValue(newValue As String, row As Integer, col As Integer, Cancel As Boolean)
Public Event UpdateValue(row As Integer, col As Integer)
Public Event RolColChanged(row As Integer, col As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event ItemChange()
Public Event Click(row As Integer, col As Integer)
Public Event DblClick(row As Integer, col As Integer)

Dim oldString As String
Dim isWait_ As Boolean
Dim colEnabled_(20) As Integer
Dim colType_(20) As ColType
Dim cancelUpdate_ As Boolean
Dim autoDel_ As Boolean
Enum ColType
    TextBoxEnum = 1
    ComboBoxEnum = 2
    DateTimeEnum = 3
End Enum
Dim gridMoveType_ As Integer

Public Function getGridObj() As MSFlexGrid
    Set getGridObj = fgGrid
End Function

Private Sub cboInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        'dtInput.Visible = False
        If fgGrid.Text = "" Then
            cboInput.ListIndex = -1
        Else
            'cboInput.Text = fgGrid.Text
            SetTextToCombo fgGrid.Text, cboInput
        End If
        fgGrid.SetFocus
    ElseIf KeyCode = 13 Then
        RaiseEvent BeforeUpdateValue(cboInput.Text, fgGrid.row, fgGrid.col, cancelUpdate_)
        If cancelUpdate_ Then
            cboInput.Text = fgGrid.Text
            cancelUpdate_ = False
        Else
            fgGrid.Text = cboInput.Text
            'dtInput.Visible = False
            fgGrid.SetFocus
            SetGridPosition
            RaiseEvent UpdateValue(fgGrid.row, fgGrid.col)
        End If
    End If
End Sub

Private Sub cboInput_Validate(Cancel As Boolean)
    RaiseEvent BeforeUpdateValue(cboInput.Text, fgGrid.row, fgGrid.col, cancelUpdate_)
    If cancelUpdate_ Then
        'cboInput.ListIndex = -1
        cancelUpdate_ = False
        Cancel = True
    Else
        fgGrid.Text = cboInput.Text
        RaiseEvent UpdateValue(fgGrid.row, fgGrid.col)
    End If
End Sub

Private Sub dtInput_KeyDown(KeyCode As Integer, Shift As Integer)
Dim newStr As String
    If KeyCode = 27 Then
        'dtInput.Visible = False
        dtInput.Value = fgGrid.Text
        fgGrid.SetFocus
    ElseIf KeyCode = 13 Then
        newStr = FormatYMD_to_DMY(FormatYMD_Thai(dtInput.Day, dtInput.Month, dtInput.Year, "/"), "/", "/")
        RaiseEvent BeforeUpdateValue(Trim(newStr), fgGrid.row, fgGrid.col, cancelUpdate_)
        If cancelUpdate_ Then
            cancelUpdate_ = False
        Else
            SendKeys "{Right}"
            fgGrid.Text = Trim(newStr)
            fgGrid.SetFocus
            SetGridPosition
            RaiseEvent UpdateValue(fgGrid.row, fgGrid.col)
        End If
    End If
End Sub

Private Sub dtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub dtInput_Validate(Cancel As Boolean)
Dim newStr As String
    newStr = Trim(FormatYMD_to_DMY(FormatYMD_Thai(dtInput.Day, dtInput.Month, dtInput.Year, "/"), "/", "/"))
    RaiseEvent BeforeUpdateValue(newStr, fgGrid.row, fgGrid.col, cancelUpdate_)
    If cancelUpdate_ Then
        cancelUpdate_ = False
        Cancel = True
    Else
        fgGrid.Text = newStr
        RaiseEvent UpdateValue(fgGrid.row, fgGrid.col)
    End If
End Sub

Private Sub fgGrid_Click()
    RaiseEvent Click(fgGrid.row, fgGrid.col)
End Sub

Private Sub fgGrid_DblClick()
    If colType_(fgGrid.col) = TextBoxEnum And CBool(colEnabled_(fgGrid.col)) = True Then
        EnableTextInput
        txtInput.Text = fgGrid.Text
        txtInput.SelStart = 0
    End If
    RaiseEvent DblClick(fgGrid.row, fgGrid.col)
End Sub

Private Sub fgGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 And colType_(fgGrid.col) = TextBoxEnum Then
        EnableTextInput
        txtInput.Text = fgGrid.Text
        txtInput.SelStart = 0
    ElseIf KeyCode = 46 And autoDel_ Then
        If CBool(colEnabled_(fgGrid.col)) = True Then
            RaiseEvent BeforeUpdateValue("", fgGrid.row, fgGrid.col, cancelUpdate_)
            If cancelUpdate_ Then
                cancelUpdate_ = False
            Else
                fgGrid.Text = ""
                RaiseEvent UpdateValue(fgGrid.row, fgGrid.col)
                fgGrid.SetFocus
                SendKeys "{Right}"
            End If
        End If
    ElseIf KeyCode = 13 Then
        SetGridPosition
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub fgGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    Else
        If Not CBool(colEnabled_(fgGrid.col)) Then Exit Sub
        If colType_(fgGrid.col) = TextBoxEnum Then
            EnableTextInput
            txtInput.Text = Chr$(KeyAscii)
            txtInput.SelStart = 1
        ElseIf colType_(fgGrid.col) = DateTimeEnum Then
            dtInput.SetFocus
        ElseIf colType_(fgGrid.col) = ComboBoxEnum Then
            cboInput.SetFocus
        End If
    End If
End Sub

Private Sub fgGrid_RowColChange()
On Error GoTo getErr:
    If fgGrid.Enabled = False Then Exit Sub
    dtInput.Visible = False
    cboInput.Visible = False
    txtInput.Visible = False
    RaiseEvent RolColChanged(fgGrid.row, fgGrid.col)
    With fgGrid
        If colEnabled_(.col) = CInt(True) Then
            If colType_(.col) = ComboBoxEnum Then
                cboInput.Left = .ColPos(.col) + .Left '+ 40
                'cboInput.Height = .RowHeight(.row) - 60
                cboInput.Width = .ColWidth(.col) '- 90
                cboInput.Top = .Top + .RowPos(.row) ' + 40
                cboInput.Visible = True
                dtInput.Visible = False
                cboInput.Text = .Text
            ElseIf colType_(.col) = DateTimeEnum Then
                dtInput.Left = .ColPos(.col) + .Left '+ 40
                dtInput.Height = .RowHeight(.row) ' - 60
                dtInput.Width = .ColWidth(.col) '- 90
                dtInput.Top = .Top + .RowPos(.row) ' + 40
                If IsDate(.TextMatrix(.row, .col)) Then
                    dtInput.Value = .TextMatrix(.row, .col)
                Else
                    dtInput.Value = Date
                End If
                dtInput.Visible = True
                cboInput.Visible = False
                'dtInput.SetFocus
            End If
        End If
    End With
    Exit Sub
getErr:
    If Err = 383 And colType_(fgGrid.col) = ComboBoxEnum Then
        cboInput.ListIndex = -1
        Exit Sub
    Else
        MsgBox "Error occured " & Err
        Exit Sub
    End If
End Sub

Private Sub fgGrid_Scroll()
fgGrid_RowColChange
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        txtInput.Text = vbNullString
        txtInput.Visible = False
        fgGrid.Text = oldString
    ElseIf KeyCode = 13 Then
        RaiseEvent BeforeUpdateValue(txtInput.Text, fgGrid.row, fgGrid.col, cancelUpdate_)
        If cancelUpdate_ Then
            txtInput.Text = oldString
            cancelUpdate_ = False
        Else
            fgGrid.Text = txtInput.Text
            txtInput.Visible = False
            fgGrid.SetFocus
            SetGridPosition
            RaiseEvent UpdateValue(fgGrid.row, fgGrid.col)
        End If
    End If
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
    RaiseEvent BeforeUpdateValue(txtInput.Text, fgGrid.row, fgGrid.col, cancelUpdate_)
    If cancelUpdate_ Then
        txtInput.Text = fgGrid.Text
        cancelUpdate_ = False
        Cancel = True
    Else
        fgGrid.Text = txtInput.Text
        RaiseEvent UpdateValue(fgGrid.row, fgGrid.col)
    End If
End Sub

Private Sub UserControl_EnterFocus()
    If fgGrid.Enabled Then fgGrid_RowColChange
    cancelUpdate_ = False
End Sub

Private Sub UserControl_ExitFocus()
Dim resValidate As Boolean
    If isWait_ Then Exit Sub
    resValidate = False
    If txtInput.Visible Then
        txtInput_Validate resValidate
        txtInput.Visible = False
    ElseIf cboInput.Visible Then
        cboInput_Validate resValidate
        cboInput.Visible = False
    ElseIf dtInput.Visible Then
        dtInput_Validate resValidate
        dtInput.Visible = False
    End If
End Sub

Private Sub UserControl_Initialize()
Dim x%
    For x = 0 To UBound(colEnabled_)
        colEnabled_(x) = CInt(True)
        colType_(x) = TextBoxEnum
    Next
    For x = 0 To fgGrid.Cols - 1
        fgGrid.FixedAlignment(x) = 4
    Next
End Sub

Private Function NextColWidth() As Long
On Error GoTo getErr:
    NextColWidth = fgGrid.ColWidth(fgGrid.col + 1)
    Exit Function
getErr:
    NextColWidth = 0
    Exit Function
End Function

Private Function NextRowHeight() As Long
On Error GoTo getErr:
    NextRowHeight = fgGrid.RowHeight(fgGrid.row + 1)
    Exit Function
getErr:
    NextRowHeight = 0
    Exit Function
End Function

Private Sub SetGridPosition()
    If gridMoveType_ = 0 Then
        If NextColWidth = 0 And NextRowHeight = 0 Then
            SendKeys "^{home}"
        ElseIf NextRowHeight = 0 Then
            SendKeys "{Right}"
            SendKeys "^{Up}"
        Else
            SendKeys "{Down}"
        End If
    ElseIf gridMoveType_ = 1 Then
        If NextColWidth = 0 And NextRowHeight = 0 Then
            SendKeys "^{home}"
        ElseIf NextColWidth = 0 Then
            SendKeys "{Down}"
            SendKeys "^{Left}"
'        ElseIf NextRowHeight = 0 Then
'            SendKeys "{Right}"
'            SendKeys "^{Up}"
        Else
            SendKeys "{Right}"
        End If
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Public Function isEmpty() As Boolean
    If fgGrid.Rows > 1 Then
        If Trim(fgGrid.TextMatrix(1, 0)) = vbNullString Then
            isEmpty = True
        Else
            isEmpty = False
        End If
    Else
        isEmpty = True
    End If
End Function

Public Sub ClearAllData()
    With fgGrid
        .Rows = 1
        .Rows = 2
    End With
End Sub

Public Property Get ColEnabled(col As Integer) As Boolean
    ColEnabled(col) = CBool(colEnabled_(col))
End Property

Public Property Let ColEnabled(col As Integer, ByVal newval As Boolean)
    colEnabled_(col) = CInt(newval)
End Property

Private Sub UserControl_Resize()
    fgGrid.Width = UserControl.Width
    fgGrid.Height = UserControl.Height
End Sub

Public Property Get ColInputType(selCol As Integer) As ColType
    ColInputType = colType_(selCol)
End Property

Public Property Let ColInputType(selCol As Integer, ByVal newval As ColType)
    colType_(selCol) = newval
End Property

Public Function getColObject(selCol As Integer) As Object
    If colType_(selCol) = ComboBoxEnum Then
        Set getColObject = cboInput
    ElseIf colType_(selCol) = DateTimeEnum Then
        Set getColObject = dtInput
    ElseIf colType_(selCol) = TextBoxEnum Then
        Set getColObject = txtInput
    End If
End Function

Private Sub EnableTextInput()
    
            With fgGrid
                If .ColWidth(.col) < 90 Then Exit Sub
                If .RowHeight(.row) < 70 Then Exit Sub
                oldString = .Text
                txtInput.Font.Name = .Font.Name
                txtInput.Font.Size = .Font.Size
                If .CellForeColor <> 0 Then
                    txtInput.ForeColor = .CellForeColor
                Else
                    txtInput.ForeColor = vbBlack
                End If
                If .CellBackColor <> 0 Then
                    txtInput.BackColor = .CellBackColor
                Else
                    txtInput.BackColor = vbWhite
                End If
                txtInput.Left = .CellLeft + 40 ' .ColPos(.col) + .Left + 40
                txtInput.Height = .RowHeight(.row) - 70
                txtInput.Width = .ColWidth(.col) - 90
                txtInput.Top = .Top + .RowPos(.row) + 60
                If .ColAlignment(.col) = 7 Then
                    txtInput.Alignment = 1
                Else
                    txtInput.Alignment = 0
                End If
                txtInput.Visible = True
                txtInput.SetFocus
            End With
End Sub

Public Sub ClearRowData(selRow As Integer)
Dim x%
    With fgGrid
        For x = 0 To .Cols - 1
            .TextMatrix(selRow, x) = vbNullString
        Next
    End With
End Sub

Public Property Get AutoDelete() As Boolean
    AutoDelete = autoDel_
End Property
Public Property Let AutoDelete(ByVal newval As Boolean)
    autoDel_ = newval
End Property

Public Sub GridMoveType(moveType As Integer)
    gridMoveType_ = moveType
End Sub

Public Sub WaitLostFocust(isWait As Boolean)
    isWait_ = isWait
End Sub

