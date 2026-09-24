VERSION 5.00
Begin VB.UserControl ctlDate 
   BackStyle       =   0  'Transparent
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1800
   KeyPreview      =   -1  'True
   ScaleHeight     =   315
   ScaleWidth      =   1800
   Begin VB.TextBox txtYear 
      Height          =   315
      Left            =   1260
      MaxLength       =   4
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.ComboBox cboMonth 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   450
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   675
   End
   Begin VB.TextBox txtDay 
      Height          =   315
      Left            =   0
      MaxLength       =   2
      TabIndex        =   0
      Top             =   0
      Width           =   315
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   330
      X2              =   420
      Y1              =   240
      Y2              =   30
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1140
      X2              =   1230
      Y1              =   240
      Y2              =   30
   End
End
Attribute VB_Name = "ctlDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Enum LocalType
    Thai = 1
    English = 2
End Enum
Enum YearType
    Thi = 1
    Eng = 2
End Enum
Dim yearType_ As YearType
Dim localType_ As LocalType
Dim day_ As Integer
Dim month_ As Integer
Dim year_ As Integer
Dim bgColor_ As OLE_COLOR

Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event DateChanged()

Private Sub cboMonth_Click()
    month_ = cboMonth.ListIndex + 1
    RaiseEvent DateChanged
End Sub

Private Sub txtDay_GotFocus()
    txtDay.SelStart = 0
    txtDay.SelLength = Len(txtDay.Text)
End Sub

Private Sub txtDay_KeyUp(KeyCode As Integer, Shift As Integer)
    If IsNumeric(txtDay) And Trim(txtDay) <> "" Then
        day_ = CInt(txtDay.Text)
    End If
    RaiseEvent DateChanged
End Sub

Private Sub txtDay_Validate(Cancel As Boolean)
    If Trim(txtDay.Text) = "" Or IsNumeric(txtDay.Text) = False Then
        MsgBox "ข้อมูลวันที่ไม่ถูกต้อง", vbExclamation
        txtDay.Text = day_
        Cancel = True
    Else
        day_ = CInt(txtDay.Text)
    End If
End Sub

Private Sub txtYear_GotFocus()
    txtYear.SelStart = 0
    txtYear.SelLength = Len(txtYear.Text)
End Sub

Private Sub txtYear_KeyUp(KeyCode As Integer, Shift As Integer)
    If IsNumeric(txtYear) And Trim(txtYear) <> "" Then
        year_ = CInt(txtYear.Text)
    End If
    RaiseEvent DateChanged
End Sub

Private Sub txtYear_Validate(Cancel As Boolean)
    If Trim(txtYear.Text) = "" Or IsNumeric(txtYear.Text) = False Then
        MsgBox "ข้อมูลปีไม่ถูกต้อง", vbExclamation
        txtYear.Text = year_
        Cancel = True
    Else
        year_ = CheckYearValue(CInt(txtYear.Text))
    End If
End Sub

Private Sub UserControl_ExitFocus()
    txtDay_Validate False
    txtYear_Validate False
End Sub

Private Sub UserControl_Initialize()
    cboMonth.Clear
    With cboMonth
        .AddItem "ม.ค."
        .AddItem "ก.พ."
        .AddItem "มี.ค."
        .AddItem "เม.ย."
        .AddItem "พ.ค."
        .AddItem "มิ.ย."
        .AddItem "ก.ค."
        .AddItem "ส.ค."
        .AddItem "ก.ย."
        .AddItem "ต.ค."
        .AddItem "พ.ย."
        .AddItem "ธ.ค."
    End With
    DayValue = Day(Date)
    MonthValue = Month(Date)
    YearValue = Year(Date)
End Sub


Public Property Get YearType() As YearType
    YearType = yearType_
End Property
Public Property Let YearType(ByVal newVal As YearType)
    yearType_ = newVal
End Property

Public Property Get LocalType() As YearType
    LocalType = localType_
End Property
Public Property Let LocalType(ByVal newVal As YearType)
    localType_ = newVal
End Property

Public Property Get DayValue() As Integer
    DayValue = day_
End Property
Public Property Let DayValue(ByVal newVal As Integer)
    day_ = newVal
    txtDay.Text = Format(day_, "00")
End Property

Public Property Get MonthValue() As Integer
    MonthValue = month_
End Property
Public Property Let MonthValue(ByVal newVal As Integer)
    month_ = newVal
    cboMonth.ListIndex = month_ - 1
End Property

Public Property Get YearValue() As Integer
    YearValue = year_
End Property
Public Property Let YearValue(ByVal newVal As Integer)
    year_ = CheckYearValue(newVal)
    txtYear.Text = year_
End Property

Public Property Get ValueDMY() As String
    ValueDMY = Format(day_, "00") & "/" & Format(month_, "00") & "/" & year_
End Property
Public Property Let ValueDMY(ByVal newVal As String)
Dim tmpArr
Dim tmpStr As String
    tmpArr = Split(newVal, "/")
    tmpStr = tmpArr(0)
    If IsNumeric(tmpStr) And Trim(tmpStr) <> "" Then
        DayValue = CInt(tmpStr)
    End If
    
    tmpStr = tmpArr(1)
    If IsNumeric(tmpStr) And Trim(tmpStr) <> "" Then
        MonthValue = CInt(tmpStr)
    End If
    
    tmpStr = tmpArr(2)
    If IsNumeric(tmpStr) And Trim(tmpStr) <> "" Then
        YearValue = CheckYearValue(CInt(tmpStr))
    End If
End Property

Public Property Get ValueMDY() As String
    ValueMDY = Format(month_, "00") & "/" & Format(day_, "00") & "/" & year_
End Property
Public Property Let ValueMDY(ByVal newVal As String)
Dim tmpArr
Dim tmpStr As String
    tmpArr = Split(newVal, "/")
    tmpStr = tmpArr(0)
    If IsNumeric(tmpStr) And Trim(tmpStr) <> "" Then
        MonthValue = CInt(tmpStr)
    End If
    
    tmpStr = tmpArr(1)
    If IsNumeric(tmpStr) And Trim(tmpStr) <> "" Then
        DayValue = CInt(tmpStr)
    End If
    
    tmpStr = tmpArr(2)
    If IsNumeric(tmpStr) And Trim(tmpStr) <> "" Then
        YearValue = CheckYearValue(CInt(tmpStr))
    End If
End Property

Public Property Get ValueYMD() As String
    ValueYMD = year_ & "/" & Format(month_, "00") & "/" & Format(day_, "00")
End Property
Public Property Let ValueYMD(ByVal newVal As String)
Dim tmpArr
Dim tmpStr As String
    If Trim(newVal) = "" Then Exit Property
    tmpArr = Split(newVal, "/")
    tmpStr = tmpArr(0)
    If IsNumeric(tmpStr) And Trim(tmpStr) <> "" Then
        YearValue = CheckYearValue(CInt(tmpStr))
    End If
    
    tmpStr = tmpArr(1)
    If IsNumeric(tmpStr) And Trim(tmpStr) <> "" Then
        MonthValue = CInt(tmpStr)
    End If
    
    tmpStr = tmpArr(2)
    If IsNumeric(tmpStr) And Trim(tmpStr) <> "" Then
        DayValue = CInt(tmpStr)
    End If
End Property

Private Function CheckYearValue(selYear As Integer) As Integer
    If selYear < 2500 Then
        CheckYearValue = selYear + 543
    Else
        CheckYearValue = selYear
    End If
End Function

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal newVal As OLE_COLOR)
    UserControl.BackColor = newVal
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If KeyCode = 13 Then
        txtDay_Validate False
        txtYear_Validate False
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    UserControl.BackColor = PropBag.ReadProperty("BackColor")
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", UserControl.BackColor
End Sub

Public Property Get Enabled() As Boolean
    Enabled = txtDay.Enabled
End Property
Public Property Let Enabled(ByVal newVal As Boolean)
    txtDay.Enabled = newVal
    cboMonth.Enabled = newVal
    txtYear.Enabled = newVal
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = txtDay.ForeColor
End Property
Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    txtDay.ForeColor = vNewValue
    cboMonth.ForeColor = vNewValue
    txtYear.ForeColor = vNewValue
End Property

Public Property Get MonthName() As String
        Select Case cboMonth.ListIndex
                Case 0
                    MonthName = "มกราคม"
                Case 1
                    MonthName = "กุมภาพันธ์"
                Case 2
                    MonthName = "มีนาคม"
                Case 3
                    MonthName = "เมษายน"
                Case 4
                    MonthName = "พฤษภาคม"
                Case 5
                    MonthName = "มิถุนายน"
                Case 6
                    MonthName = "กรกฎาคม"
                Case 7
                    MonthName = "สิงหาคม"
                Case 8
                    MonthName = "กันยายน"
                Case 9
                    MonthName = "ตุลาคม"
                Case 10
                    MonthName = "พฤศจิกายน"
                Case 11
                    MonthName = "ธันวาคม"
        End Select
End Property

Public Property Get MonthNameShort() As String
        Select Case cboMonth.ListIndex
                Case 0
                    MonthNameShort = "ม.ค."
                Case 1
                    MonthNameShort = "ก.พ."
                Case 2
                    MonthNameShort = "มี.ค."
                Case 3
                    MonthNameShort = "เม.ย."
                Case 4
                    MonthNameShort = "พ.ค."
                Case 5
                    MonthNameShort = "มิ.ย."
                Case 6
                    MonthNameShort = "ก.ค."
                Case 7
                    MonthNameShort = "ส.ค."
                Case 8
                    MonthNameShort = "ก.ย."
                Case 9
                    MonthNameShort = "ต.ค."
                Case 10
                    MonthNameShort = "พ.ย."
                Case 11
                    MonthNameShort = "ธ.ค."
        End Select
End Property

