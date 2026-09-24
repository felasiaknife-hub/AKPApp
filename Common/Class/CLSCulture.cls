VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLSCulture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()

Private Const LOCALE_ICENTURY = &H24        '  century format specifier
Private Const LOCALE_ICOUNTRY = &H5         '  country code
Private Const LOCALE_ICURRDIGITS = &H19        '  # local monetary digits
Private Const LOCALE_ICURRENCY = &H1B        '  positive currency mode
Private Const LOCALE_IDATE = &H21        '  short date format ordering
Private Const LOCALE_IDAYLZERO = &H26        '  leading zeros in day field
Private Const LOCALE_IDEFAULTCODEPAGE = &HB         '  default code page
Private Const LOCALE_IDEFAULTCOUNTRY = &HA         '  default country code
Private Const LOCALE_IDEFAULTLANGUAGE = &H9         '  default language id
Private Const LOCALE_IDIGITS = &H11        '  number of fractional digits
Private Const LOCALE_IINTLCURRDIGITS = &H1A        '  # intl monetary digits
Private Const LOCALE_ILANGUAGE = &H1         '  language id
Private Const LOCALE_ILDATE = &H22        '  long date format ordering
Private Const LOCALE_ILZERO = &H12        '  leading zeros for decimal
Private Const LOCALE_IMEASURE = &HD         '  0 = metric, 1 = US
Private Const LOCALE_IMONLZERO = &H27        '  leading zeros in month field
Private Const LOCALE_INEGCURR = &H1C        '  negative currency mode
Private Const LOCALE_INEGSEPBYSPACE = &H57        '  mon sym sep by space from neg amt
Private Const LOCALE_INEGSIGNPOSN = &H53        '  negative sign position
Private Const LOCALE_INEGSYMPRECEDES = &H56        '  mon sym precedes neg amt
Private Const LOCALE_IPOSSEPBYSPACE = &H55        '  mon sym sep by space from pos amt
Private Const LOCALE_IPOSSIGNPOSN = &H52        '  positive sign position
Private Const LOCALE_IPOSSYMPRECEDES = &H54        '  mon sym precedes pos amt
Private Const LOCALE_ITIME = &H23        '  time format specifier
Private Const LOCALE_ITLZERO = &H25        '  leading zeros in time field
Private Const LOCALE_NOUSEROVERRIDE = &H80000000  '  do not use user overrides
Private Const LOCALE_S1159 = &H28        '  AM designator
Private Const LOCALE_S2359 = &H29        '  PM designator
Private Const LOCALE_SABBREVCTRYNAME = &H7         '  abbreviated country name
Private Const LOCALE_SABBREVDAYNAME1 = &H31        '  abbreviated name for Monday
Private Const LOCALE_SABBREVDAYNAME3 = &H33        '  abbreviated name for Wednesday
Private Const LOCALE_SABBREVDAYNAME2 = &H32        '  abbreviated name for Tuesday
Private Const LOCALE_SABBREVDAYNAME4 = &H34        '  abbreviated name for Thursday
Private Const LOCALE_SABBREVDAYNAME5 = &H35        '  abbreviated name for Friday
Private Const LOCALE_SABBREVDAYNAME6 = &H36        '  abbreviated name for Saturday
Private Const LOCALE_SABBREVDAYNAME7 = &H37        '  abbreviated name for Sunday
Private Const LOCALE_SABBREVLANGNAME = &H3         '  abbreviated language name
Private Const LOCALE_SABBREVMONTHNAME1 = &H44        '  abbreviated name for January
Private Const LOCALE_SABBREVMONTHNAME10 = &H4D        '  abbreviated name for October
Private Const LOCALE_SABBREVMONTHNAME11 = &H4E        '  abbreviated name for November
Private Const LOCALE_SABBREVMONTHNAME12 = &H4F        '  abbreviated name for December
Private Const LOCALE_SABBREVMONTHNAME13 = &H100F
Private Const LOCALE_SABBREVMONTHNAME2 = &H45        '  abbreviated name for February
Private Const LOCALE_SABBREVMONTHNAME3 = &H46        '  abbreviated name for March
Private Const LOCALE_SABBREVMONTHNAME4 = &H47        '  abbreviated name for April
Private Const LOCALE_SABBREVMONTHNAME5 = &H48        '  abbreviated name for May
Private Const LOCALE_SABBREVMONTHNAME6 = &H49        '  abbreviated name for June
Private Const LOCALE_SABBREVMONTHNAME7 = &H4A        '  abbreviated name for July
Private Const LOCALE_SABBREVMONTHNAME8 = &H4B        '  abbreviated name for August
Private Const LOCALE_SABBREVMONTHNAME9 = &H4C        '  abbreviated name for September
Private Const LOCALE_SCOUNTRY = &H6         '  localized name of country
Private Const LOCALE_SCURRENCY = &H14        '  local monetary symbol
Private Const LOCALE_SDATE = &H1D        '  date separator
Private Const LOCALE_SDAYNAME1 = &H2A        '  long name for Monday
Private Const LOCALE_SDAYNAME2 = &H2B        '  long name for Tuesday
Private Const LOCALE_SDAYNAME3 = &H2C        '  long name for Wednesday
Private Const LOCALE_SDAYNAME4 = &H2D        '  long name for Thursday
Private Const LOCALE_SDAYNAME5 = &H2E        '  long name for Friday
Private Const LOCALE_SDAYNAME6 = &H2F        '  long name for Saturday
Private Const LOCALE_SDAYNAME7 = &H30        '  long name for Sunday
Private Const LOCALE_SDECIMAL = &HE         '  decimal separator
Private Const LOCALE_SENGCOUNTRY = &H1002      '  English name of country
Private Const LOCALE_SENGLANGUAGE = &H1001      '  English name of language
Private Const LOCALE_SGROUPING = &H10        '  digit grouping
Private Const LOCALE_SINTLSYMBOL = &H15        '  intl monetary symbol
Private Const LOCALE_SLANGUAGE = &H2         '  localized name of language
Private Const LOCALE_SLIST = &HC         '  list item separator
Private Const LOCALE_SLONGDATE = &H20        '  long date format string
Private Const LOCALE_SMONDECIMALSEP = &H16        '  monetary decimal separator
Private Const LOCALE_SMONGROUPING = &H18        '  monetary grouping
Private Const LOCALE_SMONTHNAME1 = &H38        '  long name for January
Private Const LOCALE_SMONTHNAME10 = &H41        '  long name for October
Private Const LOCALE_SMONTHNAME11 = &H42        '  long name for November
Private Const LOCALE_SMONTHNAME12 = &H43        '  long name for December
Private Const LOCALE_SMONTHNAME2 = &H39        '  long name for February
Private Const LOCALE_SMONTHNAME3 = &H3A        '  long name for March
Private Const LOCALE_SMONTHNAME5 = &H3C        '  long name for May
Private Const LOCALE_SMONTHNAME6 = &H3D        '  long name for June
Private Const LOCALE_SMONTHNAME7 = &H3E        '  long name for July
Private Const LOCALE_SMONTHNAME8 = &H3F        '  long name for August
Private Const LOCALE_SMONTHNAME9 = &H40        '  long name for September
Private Const LOCALE_SMONTHOUSANDSEP = &H17        '  monetary thousand separator
Private Const LOCALE_SNATIVECTRYNAME = &H8         '  native name of country
Private Const LOCALE_SNATIVEDIGITS = &H13        '  native ascii 0-9
Private Const LOCALE_SNATIVELANGNAME = &H4         '  native name of language
Private Const LOCALE_SNEGATIVESIGN = &H51        '  negative sign
Private Const LOCALE_SPOSITIVESIGN = &H50        '  positive sign
Private Const LOCALE_SSHORTDATE = &H1F        '  short date format string
Private Const LOCALE_STHOUSAND = &HF         '  thousand separator
Private Const LOCALE_STIME = &H1E        '  time separator
Private Const LOCALE_SSHORTTIME = &H1003 ' short time format string
Private Const LOCALE_STIMEFORMAT = &H1003      '  time format string

Public Enum EnumLangue
    LANG_DEFAULT = &H0
    LANG_AFRIKAANS = &H36
    LANG_ALBANIAN = &H1C
    LANG_ARABIC = &H1
    LANG_ARMENIAN = &H2B
    LANG_ASSAMESE = &H4D
    LANG_AZERI = &H2C
    LANG_BASQUE = &H2D
    LANG_BELARUSIAN = &H23
    LANG_BENGALI = &H45
    LANG_BULGARIAN = &H2
    LANG_CATALAN = &H3
    LANG_CHINESE = &H4
    LANG_CROATIAN = &H1A
    LANG_CZECH = &H5
    LANG_DANISH = &H6
    LANG_DUTCH = &H13
    LANG_ENGLISH = &H9
    LANG_ESTONIAN = &H25
    LANG_FAEROESE = &H38
    LANG_FARSI = &H29
    LANG_FINNISH = &HB
    LANG_FRENCH = &HC
    LANG_GEORGIAN = &H37
    LANG_GERMAN = &H7
    LANG_GREEK = &H8
    LANG_GUJARATI = &H47
    LANG_HEBREW = &HD
    LANG_HINDI = &H39
    LANG_HUNGARIAN = &HE
    LANG_ICELANDIC = &HF
    LANG_INDONESIAN = &H21
    LANG_ITALIAN = &H10
    LANG_JAPANESE = &H11
    LANG_KANNADA = &H4B
    LANG_KASHMIRI = &H60
    LANG_KAZAK = &H3F
    LANG_KONKANI = &H57
    LANG_KOREAN = &H12
    LANG_LATVIAN = &H26
    LANG_LITHUANIAN = &H27
    LANG_MACEDONIAN = &H2F
    LANG_MALAY = &H3E
    LANG_MALAYALAM = &H4C
    LANG_MANIPURI = &H58
    LANG_MARATHI = &H4E
    LANG_NEPALI = &H61
    LANG_NORWEGIAN = &H14
    LANG_ORIYA = &H48
    LANG_POLISH = &H15
    LANG_PORTUGUESE = &H16
    LANG_PUNJABI = &H46
    LANG_ROMANIAN = &H18
    LANG_RUSSIAN = &H19
    LANG_SANSKRIT = &H4F
    LANG_SERBIAN = &H1A
    LANG_SINDHI = &H59
    LANG_SLOVAK = &H1B
    LANG_SLOVENIAN = &H24
    LANG_SPANISH = &HA
    LANG_SWAHILI = &H41
    LANG_SWEDISH = &H1D
    LANG_TAMIL = &H49
    LANG_TATAR = &H44
    LANG_TELUGU = &H4A
    ' =========================
    LANG_THAI = &H1E
    ' =========================
    LANG_TURKISH = &H1F
    LANG_UKRAINIAN = &H22
    LANG_URDU = &H20
    LANG_UZBEK = &H43
    LANG_VIETNAMESE = &H2A
    LANG_YIDDISH = &H3D
End Enum

Private Declare Function GetKeyboardLayoutName Lib "user32" Alias _
    "GetKeyboardLayoutNameA" ( _
    ByVal pwszKLID As String _
    ) As Long
    
Private Declare Function LoadKeyboardLayout Lib "user32" Alias _
    "LoadKeyboardLayoutA" ( _
    ByVal pwszKLID As String, _
    ByVal Flags As Long _
    ) As Long
    
Private Declare Function PostMessage _
    Lib "user32" Alias "PostMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
' กำหนดค่าคงที่ Flags ในการสั่งให้ LoadKeyboardLayout  ทำงาน (Activate)
Private Const KLF_ACTIVATE = &H1
Private Const STR_LANG_ENGLISH As String = "English (United States)"
Private Const STR_LANG_THAI As String = "Thai"

Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_SETTINGCHANGE = &H1A

Private Const HKEY_CURRENT_USER = &H80000001
Private sLocal As String
Private Const REG_SZ As Long = 1
Private hKeyHandle As Long

Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful..
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

Public Function GetRegionalSettings() As String  ' Retrieve the regional setting

      Dim Symbol As String
      Dim iRet1 As Long
      Dim iRet2 As Long
      Dim lpLCDataVar As String
      Dim pos As Integer
'      Dim LCID As Long
       
      LCID = GetUserDefaultLCID()

'LOCALE_SDATE is the constant for the date separator
'as stated in declarations
'for any other locale setting just change the constant

'Function can also be re-written to take the
'locale symbol being requested as a parameter
       
      iRet1 = GetLocaleInfo(LCID, LOCALE_SDATE, _
      lpLCDataVar, 0)
      Symbol = String$(iRet1, 0)
       
      iRet2 = GetLocaleInfo(LCID, LOCALE_SDATE, Symbol, iRet1)
      pos = InStr(Symbol, Chr$(0))
      If pos > 0 Then
           Symbol = Left$(Symbol, pos - 1)
'           MsgBox "Regional Setting = " + Symbol
      End If
    GetRegionalSettings = LCID
End Function

Public Sub SetRegionalSettings(ByVal fLCID As Long, ByVal CType As Long, ByVal InSymbol As String, Optional ByVal InShortTime As String = "HH:mm", Optional ByVal InLongTime As String = "HH:mm:ss")  'Change the regional setting
    Dim Symbol As String
    Dim iRet As Long
    Dim success As Long
'    Dim LCID As Long

'    LCID = GetUserDefaultLCID() 'Get user Locale ID
    LCID = fLCID
'    MsgBox GetUserLocaleInfo(LCID, LOCALE_IDEFAULTLANGUAGE)
'    Symbol = "/" 'New character for the locale
'    iRet = SetLocaleInfo(Locale, LOCALE_SDATE, Symbol)
'    If LCID = 1054 Then

''        If SetLocaleInfo(LCID, LOCALE_SLANGUAGE, "041e") = False Then MsgBox "Set language failed."
''        SwitchKeyboardLang STR_LANG_THAI
''    Else
''        If SetLocaleInfo(LCID, LOCALE_SLANGUAGE, "0409") = False Then MsgBox "Set language failed."
''        SwitchKeyboardLang STR_LANG_ENGLISH
'    End If
    
'    LCID = GetSystemDefaultLCID()
    If LCID = 0 Then Exit Sub
    DefaultLocale = GetUserLocaleInfo(DefaultLCID, LOCALE_ILANGUAGE)
    
    'localized name of language
    If GetUserLocaleInfo(DefaultLCID, LOCALE_SLANGUAGE) = "Thai" Then
    
        'localized name of calendar
    '    Debug.Print GetLocaleCalendarName(DefaultLCID)
        
        'enumerated system long date formats
    '    Call GetLocaleDateFormats(LCID, CAL_SLONGDATE)
        
        'enumerated system short date formats
    '    Debug.Print GetLocaleDateFormats(LCID, CAL_SSHORTDATE)
    
        Symbol = GetUserLocaleInfo(DefaultLCID, LOCALE_ICALENDARTYPE)
        DefaultCalendarType = Symbol
        Symbol = CType '** "7"
        If SetLocaleInfo(DefaultLCID, LOCALE_ICALENDARTYPE, Symbol) = False Then MsgBox "Set Calendar type is failed."
    Else
'        Debug.Print GetUserLocaleInfo(DefaultLCID, LOCALE_ILANGUAGE)
        '*** Set Locale to Thai
        'create reg key to add a system Intenational
        'if it already exists, then RegCreateKey just opens it.
        sLocal = "0000041e"
'        iRet = SetLocaleInfo(LCID, LOCALE_ILANGUAGE, sLocal)
        iRet = RegCreateKey(HKEY_CURRENT_USER, "Control Panel\International\", hKeyHandle)
        iRet = RegSetValueEx(hKeyHandle, "Locale", 0&, REG_SZ, ByVal sLocal, Len(sLocal))
        '****************
        If hKeyHandle Then
            Symbol = GetUserLocaleInfo(DefaultLCID, LOCALE_ICALENDARTYPE)
            DefaultCalendarType = Symbol
            Symbol = CType '** "7"
            iRet = SetLocaleInfo(LCID, LOCALE_ICALENDARTYPE, Symbol)
        End If
    End If
    
    Symbol = GetUserLocaleInfo(LCID, LOCALE_SSHORTDATE)
    DefaultShortDate = Symbol
    Symbol = InSymbol
    iRet = SetLocaleInfo(LCID, LOCALE_SSHORTDATE, Symbol)  ' MsgBox "Set Date format failed."
    
    Symbol = GetUserLocaleInfo(LCID, LOCALE_SLONGDATE)
    DefaultLongDate = Symbol
    Symbol = "d MMMM yyyy"
    iRet = SetLocaleInfo(LCID, LOCALE_SLONGDATE, Symbol)  '= False Then MsgBox "Set long date format failed."
    
    Symbol = GetUserLocaleInfo(LCID, LOCALE_SSHORTTIME)
    DefaultShortTime = Symbol
    Symbol = "HH:mm:ss"
    iRet = SetLocaleInfo(LCID, LOCALE_SSHORTTIME, Symbol) '= False Then MsgBox "Set short time format failed."
    
    PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
End Sub

Public Sub SetDefaultRegionalSettings(ByVal fLCID As Long)  'Change the regional setting
    Dim Symbol As String
    Dim iRet As Long
'    Dim LCID As Long
    Dim success As Long
    
    LCID = fLCID

    If LCID = 0 Then Exit Sub
        
    '*** Set Locale Default
    'create reg key to add a system Intenational
    'if it already exists, then RegCreateKey just opens it.
    sLocal = Format(DefaultLocale, "00000000")
    iRet = RegCreateKey(HKEY_CURRENT_USER, "Control Panel\International\", hKeyHandle)
    iRet = RegSetValueEx(hKeyHandle, "Locale", 0&, REG_SZ, ByVal sLocal, Len(sLocal))
        
    iRet = SetLocaleInfo(LCID, LOCALE_SLANGUAGE, DefaultLCID)
    PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
    
    Symbol = DefaultCalendarType
    iRet = SetLocaleInfo(DefaultLCID, LOCALE_ICALENDARTYPE, Symbol) '= False Then MsgBox "Set Calendar type is failed."
    
    Symbol = DefaultShortDate
    iRet = SetLocaleInfo(LCID, LOCALE_SSHORTDATE, Symbol)  '= False Then MsgBox "Set Date format failed."
    
    Symbol = DefaultLongDate
    iRet = SetLocaleInfo(LCID, LOCALE_SLONGDATE, Symbol)  '= False Then MsgBox "Set long date format failed."
    
    Symbol = DefaultShortTime
    iRet = SetLocaleInfo(LCID, LOCALE_SSHORTTIME, Symbol) '= False Then MsgBox "Set short time format failed."
    
    PostMessage HWND_BROADCAST, WM_SETTINGCHANGE, 0, 0
End Sub

Public Function SwitchKeyboardLang(ByVal strLangID As String) As Boolean
Dim strRet As String
On Error Resume Next

    ' ----------------------------------------------------------------------------------------------------------------
    ' แท้ที่จริงแล้ว เราสั่งด้วยคำสั่งบรรทัดเดียวก็เป็นอันจบแล้วแหละครับ
    ' strRet = LoadKeyboardLayout(strLangID, KLF_ACTIVATE)
    ' ผมถ่ายทอดแนวคิดในแบบฉบับนี้ของผมออกไป ก็เพื่อต้องการให้พี่น้อง ... ได้คิด ได้ฝึกฝน
    ' และต้องรอบคอบ ระมัดระวัง ต่อการใช้งาน Win32 API ให้ดีด้วย
    ' ----------------------------------------------------------------------------------------------------------------
    
    ' คือ จองพื้นที่แบบ String เอาไว้ 8 ตัว ตามจำนวนของ Locale identifier
    ' ลองคุณไม่จองเอาไว้ก่อนดูครับว่าจะมีผลเช่นไร ... คิดๆๆๆๆ แล้ว ก็ทำๆๆๆๆ
    strRet = Space$(8)
    ' หรือ strRet = String$(8, 0) ... จองพื้นที่ว่างเปล่าใน Memory 8 ไบต์ (เมื่อ ASCII 0 = Null)
    
    ' อ่านค่ารหัสภาษา (Locale identifier) สถานะปัจจุบันก่อน
    GetKeyboardLayoutName strRet
    
    ' หากภาษาที่กำหนดมามันตรงกับของเดิมอยู่แล้ว ก็ไม่ต้องไปทำอะไรอีก
    ' ส่งค่า TRUE กลับออกไปได้เลย ... เพื่อความรวดเร็ว และไม่เสียเวลา
    If strRet = strLangID Then
        SwitchKeyboardLang = True
        Exit Function
    Else
        ' เคลียร์ค่าใหม่
        strRet = Space$(8)
        ' ติดต่อกับระบบปฏิบัติการ (OS) เพื่อให้โหลดรูปแบบภาษาตามที่กำหนดเอาไว้
        strRet = LoadKeyboardLayout(strLangID, KLF_ACTIVATE)
    End If
    
    ' ทดสอบหากเปลี่ยนภาษาได้สำเร็จ โดยอ่านค่าสถานะปัจจุบันเข้ามาอีกรอบ
    ' หรืออีกแง่มุมหนึ่ง ก็คือ การนำไปทดสอบว่าสถานะปัจจุบัน แป้นคีย์บอร์ดมันเป็นภาษาอะไร
    GetKeyboardLayoutName strRet
    
    ' หากโหลดภาษาได้ตามที่ต้องการให้ส่งค่า TRUE กลับไป (DTAC ไม่ต้อง เพราะผมยังใช้งานอยู่ 55555+)
    If strRet = strLangID Then
        SwitchKeyboardLang = True
    Else
        SwitchKeyboardLang = False
    End If
End Function
