Attribute VB_Name = "modCalendarInfo"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ฉ1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Public thisCombo As ComboBox

Public Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, _
   ByVal LCType As Long, _
   ByVal lpLCData As String, _
   ByVal cchData As Long) As Long
   
Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" ( _
    ByVal Locale As Long, _
    ByVal LCType As Long, _
    ByVal lpLCData As String) _
    As Boolean
    
Public Declare Function EnumCalendarInfo Lib "kernel32" _
   Alias "EnumCalendarInfoA" _
  (ByVal lpCalInfoEnumProc As Long, _
   ByVal Locale As Long, _
   ByVal Calendar As Long, _
   ByVal CalType As Long) As Long
   
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

Public Const LOCALE_SLANGUAGE As Long = &H2     'localized name of language

'Calendar ID Values.
Public Const CAL_GREGORIAN = 1                      'Gregorian (localized) calendar
Public Const CAL_GREGORIAN_US = 2                   'Gregorian (U.S.) calendar
Public Const CAL_JAPAN = 3                          'Japanese Emperor Era calendar
Public Const CAL_TAIWAN = 4                         'Taiwan Region Era calendar
Public Const CAL_KOREA = 5                          'Korean Tangun Era calendar
Public Const CAL_HIJRI = 6                          'Hijri (Arabic Lunar) calendar
Public Const CAL_THAI = 7                           'Thai calendar
Public Const CAL_HEBREW = 8                         'Hebrew (Lunar) calendar
Public Const CAL_GREGORIAN_ME_FRENCH = 9            'Gregorian Middle East French calendar
Public Const CAL_GREGORIAN_ARABIC = 10              'Gregorian Arabic calendar
Public Const CAL_GREGORIAN_XLIT_ENGLISH = 11        'Gregorian Transliterated English calendar
Public Const CAL_GREGORIAN_XLIT_FRENCH = 12         'Gregorian Transliterated French calendar

' Calendar Enumeration Value.
Public Const ENUM_ALL_CALENDARS As Long = &HFFFFFFFF 'enumerate all calendars
Public Const CAL_ICALINTVALUE As Long = &H1         'calendar type
Public Const CAL_SCALNAME As Long = &H2             'native name of calendar
Public Const CAL_IYEAROFFSETRANGE As Long = &H3     'starting years of eras
Public Const CAL_SERASTRING As Long = &H4           'era name for IYearOffsetRanges
Public Const CAL_SSHORTDATE As Long = &H1F           'short date format string
Public Const CAL_SLONGDATE As Long = &H6            'long date format string
Public Const CAL_SDAYNAME1 As Long = &H7            'native name for Monday
Public Const CAL_SDAYNAME2 As Long = &H8            'native name for Tuesday
Public Const CAL_SDAYNAME3 As Long = &H9            'native name for Wednesday
Public Const CAL_SDAYNAME4 As Long = &HA            'native name for Thursday
Public Const CAL_SDAYNAME5 As Long = &HB            'native name for Friday
Public Const CAL_SDAYNAME6 As Long = &HC            'native name for Saturday
Public Const CAL_SDAYNAME7 As Long = &HD            'native name for Sunday
Public Const CAL_SABBREVDAYNAME1 As Long = &HE      'abbreviated name for Monday
Public Const CAL_SABBREVDAYNAME2 As Long = &HF      'abbreviated name for Tuesday
Public Const CAL_SABBREVDAYNAME3 As Long = &H10     'abbreviated name for Wednesday
Public Const CAL_SABBREVDAYNAME4 As Long = &H11     'abbreviated name for Thursday
Public Const CAL_SABBREVDAYNAME5 As Long = &H12     'abbreviated name for Friday
Public Const CAL_SABBREVDAYNAME6 As Long = &H13     'abbreviated name for Saturday
Public Const CAL_SABBREVDAYNAME7 As Long = &H14     'abbreviated name for Sunday
Public Const CAL_SMONTHNAME1 As Long = &H15         'native name for January
Public Const CAL_SMONTHNAME2 As Long = &H16         'native name for February
Public Const CAL_SMONTHNAME3 As Long = &H17         'native name for March
Public Const CAL_SMONTHNAME4 As Long = &H18         'native name for April
Public Const CAL_SMONTHNAME5 As Long = &H19         'native name for May
Public Const CAL_SMONTHNAME6 As Long = &H1A         'native name for June
Public Const CAL_SMONTHNAME7 As Long = &H1B         'native name for July
Public Const CAL_SMONTHNAME8 As Long = &H1C         'native name for August
Public Const CAL_SMONTHNAME9 As Long = &H1D         'native name for September
Public Const CAL_SMONTHNAME10 As Long = &H1E        'native name for October
Public Const CAL_SMONTHNAME11 As Long = &H1F        'native name for November
Public Const CAL_SMONTHNAME12 As Long = &H20        'native name for December
Public Const CAL_SMONTHNAME13 As Long = &H21        'native name for 13th month (if any)
Public Const CAL_SABBREVMONTHNAME1 As Long = &H22   'abbreviated name for January
Public Const CAL_SABBREVMONTHNAME2 As Long = &H23   'abbreviated name for February
Public Const CAL_SABBREVMONTHNAME3 As Long = &H24   'abbreviated name for March
Public Const CAL_SABBREVMONTHNAME4 As Long = &H25   'abbreviated name for April
Public Const CAL_SABBREVMONTHNAME5 As Long = &H26   'abbreviated name for May
Public Const CAL_SABBREVMONTHNAME6 As Long = &H27   'abbreviated name for June
Public Const CAL_SABBREVMONTHNAME7 As Long = &H28   'abbreviated name for July
Public Const CAL_SABBREVMONTHNAME8 As Long = &H29   'abbreviated name for August
Public Const CAL_SABBREVMONTHNAME9 As Long = &H2A   'abbreviated name for September
Public Const CAL_SABBREVMONTHNAME10 As Long = &H2B  'abbreviated name for October
Public Const CAL_SABBREVMONTHNAME11 As Long = &H2C  'abbreviated name for November
Public Const CAL_SABBREVMONTHNAME12 As Long = &H2D  'abbreviated name for December
Public Const CAL_SABBREVMONTHNAME13 As Long = &H2E  'abbreviated name for 13th month (if any)

'WINVER >= =&H0500 only
Public Const CAL_SYEARMONTH As Long = &H2F          'year month format string
Public Const LOCALE_ICALENDARTYPE   As Long = &H1009

Global LCID As Long
Global DefaultLCID As Long
Global DefaultLocale As String
Global DefaultCalendarType As String    ' เก็บค่าปฏิทินเดิมเอาไว้
Global DefaultShortDate As String             ' เก็บรูปแบบวันที่ (dd/MM/yy)
Global DefaultLongDate As String             ' เก็บรูปแบบวันที่ (dd/MM/yyyy)
Global DefaultShortTime As String
Global DefaultLongTime As String

Public Function EnumCalendarProc(lpDateFormatString As Long) As Long

  'application-defined callback function for EnumCalendarInfo
'   thisCombo.AddItem StringFromPointer(lpDateFormatString)
'   Debug.Print StringFromPointer(lpDateFormatString)
  'return 1 to continue enumeration
   EnumCalendarProc = 1
   
End Function


Public Function EnumCalendarNameProc(lpDateFormatString As Long) As String

  'application-defined callback function for EnumCalendarInfo
'   Form1.Text2.Text = StringFromPointer(lpDateFormatString)
'   Debug.Print StringFromPointer(lpDateFormatString)
  'return 1 to continue enumeration
'   EnumCalendarNameProc = 1
   EnumCalendarNameProc = StringFromPointer(lpDateFormatString)
End Function


Private Function StringFromPointer(lpString As Long) As String

   Dim pos As Long
   Dim buffer As String
   
  'pad a string to hold the data
   buffer = Space$(128)
   
  'copy the string pointed to by the return value
   CopyMemory ByVal buffer, lpString, ByVal Len(buffer)
   
  'remove the trailing null and trim
   pos = InStr(buffer, Chr$(0))
   
   If pos Then
      StringFromPointer = Left$(buffer, pos - 1)
   End If

End Function


Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, _
                                  ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful (r > 0)
   If r Then
    
     'pad the buffer with spaces
      sReturn = Space$(r)
       
     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (r > 0)
      If r Then
      
        'r holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, r - 1)
      
      End If
   
   End If
    
End Function

Public Function GetLocaleCalendarName(LCID As Long) As Long

    'enumerate the system calendar name(s)
    Dim ret As Long
  
    If LCID Then
   
        ret = EnumCalendarInfo(AddressOf EnumCalendarNameProc, _
                             LCID, _
                             CAL_THAI, _
                             CAL_SCALNAME)
    End If
   
    GetLocaleCalendarName = ret

End Function

Public Function GetLocaleDateFormats(LCID As Long, _
                                      CalType As Long) As Long

  'enumerate the system long date formats
   Dim ret As Long
  
   If LCID Then
   
     'enumerate available long date formats
      ret = EnumCalendarInfo(AddressOf EnumCalendarProc, _
                             LCID, _
                             ENUM_ALL_CALENDARS, _
                             CalType)
   End If
   
   GetLocaleDateFormats = ret
   
End Function


Public Function GetLocaleCalendarLists(LCID As Long, _
                                        TMP_CAL_START As Long, _
                                        TMP_CAL_END As Long) As Long

  'enumerate the system items for the range passed
   Dim ret As Long
   Dim item As Long
  
   If LCID Then
   
      For item = TMP_CAL_START To TMP_CAL_END
      
         ret = EnumCalendarInfo(AddressOf EnumCalendarProc, _
                                LCID, _
                                ENUM_ALL_CALENDARS, _
                                item)
                             
      Next
      
   End If
   
   GetLocaleCalendarLists = ret
   
End Function

Public Sub InitializeCalendarType()
    
'    Dim LCID As Long
'    Dim Symbol As String
'    Dim iRet As Long
'
'    ' หาค่า Locale ID (information) ในระบบ ... ไทยแลนด์ = 1054
'    LCID = GetSystemDefaultLCID()
'
'    ' ส่วนแรกจะเป็นการปรับรูปแบบปฏิทินจากอะไรก็แล้วแต่ ... ให้เป็น พุทธศักราช (พ.ศ.)
'    ' ค่าคงที่ที่ประกาศไว้ Module RegionalSetting ... ผมยกเอามาให้ดูครับ ... พี่น้อง
'    ' Public Const LOCALE_ICALENDARTYPE   As Long = &H1009
'
'    ' กระโดดไปโปรแกรมย่อย GetUserLCID เพื่อหาค่าว่าปฏิทินใช้ในรูปแบบไหนกันแน่
'    Symbol = GetUserLCID(LCID, LOCALE_ICALENDARTYPE)
'    ' ค่าที่จะถูก Return กลับมา (String) เช่น "1" หรือ "7"
'    ' --> "1" Gregorian (localized)
'    ' --> "2" Gregorian (English strings always)
'    ' --> "3" Era: Year of the Emperor (Japan)
'    ' --> "4" Era: Year of Taiwan Region
'    ' --> "5" Tangun Era(Korea)
'    ' --> "6" Hijri (Arabic lunar)
'    ' --> "7" Thai ... ไทยแลนด์มีค่าเท่ากับ 7
'    ' --> "8" Hebrew (Lunar)
'    ' --> "9" Gregorian Middle East French
'    ' --> "10" Gregorian Arabic calendar
'    ' --> "11" Gregorian Transliterated English
'    ' --> "12" Gregorian Transliterated French
'
'    ' เก็บค่าเอาไว้ในตัวแปรแบบ Global ... สำหรับตอนที่ปิดโปรแกรมแล้วก็ค่อยคืนค่ากลับไปตามเดิม
'    ' จะได้ไม่ต้องไปนั่งทะเลาะกับเจ้าของโปรแกรมอื่นที่เขาใช้อ่านปี ค.ศ.
'    DefaultCalendarType = Symbol
'
'    ' เปลี่ยนค่าปฏิทินให้เป็น พุทธศักราช
'    Symbol = CAL_THAI '** "7"
'    ' หรือทดสอบกำหนด Symbol = "1" ... ก็จะเห็นเป็นปี ค.ศ.
'    ' หากว่า iRet ถูกส่งกลับมาเป็น 0 คือ มีการเกิด Error ขึ้นมาน่ะครับ
'    iRet = SetLocaleInfo(LCID, LOCALE_ICALENDARTYPE, Symbol)
'
'    ' สำหรับรูปแบบวันที่แบบสั้น - Short Date
'    Symbol = GetUserLCID(LCID, LOCALE_SSHORTDATE)
'    ' เก็บค่าเดิม ... เพื่อไม่ต้องการไปทะเลาะกับใคร ... 55555+
'    DefaultShortDate = Symbol
'    ' เปลี่ยนค่าตามที่ต้องการ ... M ตัวใหญ่น่ะครับ
'    Symbol = "dd/MM/yy"
'    iRet = SetLocaleInfo(LCID, LOCALE_SSHORTDATE, Symbol)
'
'
'    ' สำหรับรูปแบบวันที่แบบยาว -  Long Date
'    Symbol = GetUserLCID(LCID, LOCALE_SLONGDATE)
'    DefaultLongDate = Symbol
'    ' เปลี่ยนค่าตามที่ต้องการ
'    Symbol = "dd/MM/yyyy"
'    iRet = SetLocaleInfo(LCID, LOCALE_SLONGDATE, Symbol)

End Sub

Public Function GetUserLCID(ByRef LCID As Long, ByVal LCType As Long) As String

    Dim Symbol As String
    Dim iRet As Long

    ' หาขนาดของชนิดข้อมูลที่ต้องการ (Locale Type) เช่น ...
    ' กรณีของรูปแบบปฏิทิน LOCALE_ICALENDARTYPE = &H1009  หรือ 4105 (ฐาน 10)
    ' กรณีของวันที่แบบสั้น LOCALE_SSHORTDATE = &H1F หรือ 31 (ฐาน 10)
    ' ค่าที่ว่ามานี้ เป็นการกำหนดค่าคงที่เอาไว้ เพื่อใช้งานผ่านทาง WinAPI น่ะครับผม
    iRet = GetLocaleInfo(LCID, LCType, Symbol, Len(Symbol))
    
    ' ไม่มีการเกิด Error ... ค่า iRet จะไม่เป็น 0 (หากเป็น 0 หรือเป็น เท็จ หรือ False นั่นเองครับ ... จำไว้ให้ดี)
    If iRet Then
        ' จับจองพื้นที่ในหน่วยความจำ (Buffer) ... Space เป็นการกำหนดขนาดตามจำนวนของ iRet
        Symbol = Space$(iRet)
        
        ' และเรียกซ้ำอีกครั้ง ... เพื่อกำหนดขนาดหน่วยจำที่ต้องการ (Buffer)
        iRet = GetLocaleInfo(LCID, LCType, Symbol, Len(Symbol))
        ' หากสามารถทำงานได้สำเร็จ ... iRet ต้องมากกว่า 0
        If iRet Then
            ' นั่นคือ iRet จะเป็นตัวกำหนดขนาดของหน่วยความจำเอาไว้
            ' คืนค่ากลับผ่านทางฟังค์ชั่น (String)
            GetUserLCID = Left$(Symbol, iRet - 1)
        End If
        
        ' จาก Statement ด้านบน เขียนสั้นแบบนี้ก็ได้ ... If iRet Then GetUserLCID = Left$(Symbol, iRet - 1)
        
    End If
    
End Function
