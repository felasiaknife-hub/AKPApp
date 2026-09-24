VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLSODBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

   Private Const REG_SZ As Long = 1
   Private Const REG_DWORD As Long = 4

   Private Const HKEY_CLASSES_ROOT = &H80000000
   Private Const HKEY_CURRENT_USER = &H80000001
   Private Const HKEY_LOCAL_MACHINE = &H80000002
   Private Const HKEY_USERS = &H80000003

   Private Const ERROR_NONE = 0
   Private Const ERROR_BADDB = 1
   Private Const ERROR_BADKEY = 2
   Private Const ERROR_CANTOPEN = 3
   Private Const ERROR_CANTREAD = 4
   Private Const ERROR_CANTWRITE = 5
   Private Const ERROR_OUTOFMEMORY = 6
   Private Const ERROR_ARENA_TRASHED = 7
   Private Const ERROR_ACCESS_DENIED = 8
   Private Const ERROR_INVALID_PARAMETERS = 87
   Private Const ERROR_NO_MORE_ITEMS = 259

   Private Const KEY_QUERY_VALUE = &H1
   Private Const KEY_SET_VALUE = &H2
   Private Const KEY_ALL_ACCESS = &H3F

   Private Const REG_OPTION_NON_VOLATILE = 0

   Private mDataSourceName As String
   Private mDatabaseName As String
   Private mDescription As String
   Private mDriverPath As String
   Private mDriverName As String
   Private mLastUser As String
   Private mUserId As String
   Private mServer As String
   Private mServerName As String
   Private mStatus As String
   
Property Get DataSourceName() As String

        DataSourceName = mDataSourceName

End Property

Property Let DataSourceName(sData As String)

        mDataSourceName = sData

End Property
   
Property Get DatabaseName() As String

        DatabaseName = mDatabaseName

End Property

Property Let DatabaseName(sData As String)

        mDatabaseName = sData

End Property
   
Property Get Description() As String

        Description = mDescription

End Property

Property Let Description(sData As String)

        mDescription = sData

End Property
   
Property Get DriverPath() As String

        DriverPath = mDriverPath

End Property

Property Let DriverPath(sData As String)

        mDriverPath = sData

End Property
   
Property Get DriverName() As String

        DriverName = mDriverName

End Property

Property Let DriverName(sData As String)
        mDriverName = sData
End Property
   
Property Get LastUser() As String

        LastUser = mLastUser

End Property

Property Let LastUser(sData As String)

        mLastUser = sData

End Property
 Property Get USERID() As String

        USERID = mUserId

End Property

Property Let USERID(sData As String)

        mUserId = sData

End Property
  
Property Get Server() As String
        Server = mServer
End Property

Property Let Server(sData As String)
        mServer = sData
End Property
Property Get Servername() As String
        Servername = mServer
End Property

Property Let Servername(sData As String)
        mServerName = sData
End Property
Property Get Status() As String

        Status = mStatus

End Property

Property Let Status(sData As String)

        mStatus = sData

End Property

Public Function SetDSN() As Boolean
   
   Dim lResult As Long
   Dim hKeyHandle As Long
   Dim Msg As String
   
   
   On Error GoTo Set_DSN_Err
   
   'create reg key to add a system DSN
   'if it already exists, then RegCreateKey just opens it.
   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & mDataSourceName, hKeyHandle)

   'Set the values of the new DSN key, or update existing ones.
   lResult = RegSetValueEx(hKeyHandle, "Database", 0&, REG_SZ, ByVal mDatabaseName, Len(mDatabaseName))
   
   lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, ByVal mDescription, Len(mDescription))
   
   lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, ByVal mDriverPath, Len(mDriverPath))
   
   lResult = RegSetValueEx(hKeyHandle, "LastUser", 0&, REG_SZ, ByVal mLastUser, Len(mLastUser))
   lResult = RegSetValueEx(hKeyHandle, "USERID", 0&, REG_SZ, ByVal mLastUser, Len(mLastUser))
   
   lResult = RegSetValueEx(hKeyHandle, "Server", 0&, REG_SZ, ByVal mServer, Len(mServer))
   lResult = RegSetValueEx(hKeyHandle, "ServerName", 0&, REG_SZ, ByVal mServer, Len(mServer))

   'Close it....
   lResult = RegCloseKey(hKeyHandle)



   'Create or Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
   
   'Specify the new value.
   lResult = RegSetValueEx(hKeyHandle, mDataSourceName, 0&, REG_SZ, ByVal mDriverName, Len(mDriverName))
   
   'Close the key.
   lResult = RegCloseKey(hKeyHandle)
   'all OK.
   SetDSN = True
   If lResult = 0 Then
        mStatus = "Succeeded!"
    Else
        mStatus = "UnSucceeded!"
    End If
   
   'later....
   Exit Function
   
Set_DSN_Err:
   
         Msg = "Error in Function clsODBC.SetDSN" & vbCrLf & vbCrLf
   Msg = Msg & "Error Number: " & Err.Number & vbCrLf & vbCrLf
   Msg = Msg & Err.Description & vbCrLf & vbCrLf
   
   mStatus = Msg
   
   SetDSN = False
   
End Function

Public Function SetUserDSN() As Boolean
   
   Dim lResult As Long
   Dim hKeyHandle As Long
   Dim Msg As String
   
   
   On Error GoTo Set_DSN_Err
   
   'create reg key to add a system DSN
   'if it already exists, then RegCreateKey just opens it.
   lResult = RegCreateKey(HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\" & mDataSourceName, hKeyHandle)

   'Set the values of the new DSN key, or update existing ones.
   lResult = RegSetValueEx(hKeyHandle, "Database", 0&, REG_SZ, ByVal mDatabaseName, Len(mDatabaseName))
   
   lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, ByVal mDescription, Len(mDescription))
   
   lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, ByVal mDriverPath, Len(mDriverPath))
   
   lResult = RegSetValueEx(hKeyHandle, "LastUser", 0&, REG_SZ, ByVal mLastUser, Len(mLastUser))
   lResult = RegSetValueEx(hKeyHandle, "USERID", 0&, REG_SZ, ByVal mLastUser, Len(mLastUser))
   
   lResult = RegSetValueEx(hKeyHandle, "Server", 0&, REG_SZ, ByVal mServer, Len(mServer))
   lResult = RegSetValueEx(hKeyHandle, "ServerName", 0&, REG_SZ, ByVal mServer, Len(mServer))

   'Close it....
   lResult = RegCloseKey(hKeyHandle)



   'Create or Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
   lResult = RegCreateKey(HKEY_CURRENT_USER, "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
   
   'Specify the new value.
   lResult = RegSetValueEx(hKeyHandle, mDataSourceName, 0&, REG_SZ, ByVal mDriverName, Len(mDriverName))
   
   'Close the key.
   lResult = RegCloseKey(hKeyHandle)
   'all OK.
   SetUserDSN = True
   If lResult = 0 Then
        mStatus = "Succeeded!"
    Else
        mStatus = "UnSucceeded!"
    End If
   
   'later....
   Exit Function
   
Set_DSN_Err:
   
    Msg = "Error in Function clsODBC.SetUserDSN" & vbCrLf & vbCrLf
    Msg = Msg & "Error Number: " & Err.Number & vbCrLf & vbCrLf
    Msg = Msg & Err.Description & vbCrLf & vbCrLf
   
    mStatus = Msg
   
    SetUserDSN = False
   
End Function

Public Function GetServer() As Boolean

   Dim lResult As Long
   Dim hKeyHandle As Long
   Dim Msg As String
      
   On Error GoTo Get_Server_Err

   lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & mDataSourceName, hKeyHandle)
   
   lResult = QueryValueEx(hKeyHandle, "Server", mServer)
   
   If lResult = ERROR_NONE Then
      GetServer = True
      mStatus = "Succeeded!"
   Else
      GetServer = False
      mStatus = "Server Not Found for this DSN!"
   End If
   
   lResult = RegCloseKey(hKeyHandle)
   

   Exit Function

Get_Server_Err:
   
         Msg = "Error in Function clsODBC.GetServer during DSN interrogation" & vbCrLf & vbCrLf
   Msg = Msg & "Error Number: " & Err.Number & vbCrLf & vbCrLf
   Msg = Msg & Err.Description & vbCrLf & vbCrLf
   
   mStatus = Msg
   
   GetServer = False

End Function
   
   Private Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
       Dim cch As Long
       Dim lrc As Long
       Dim lType As Long
       Dim lValue As Long
       Dim sValue As String

       On Error GoTo QueryValueExError

       ' Determine the size and type of data to be read
       lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
       If lrc <> ERROR_NONE Then Error 5

       Select Case lType
       
           ' For strings
           Case REG_SZ:
               sValue = String(cch, 0)
               lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
               If lrc = ERROR_NONE Then
                   vValue = Mid(sValue, 1, cch - 1)
               Else
                   vValue = Empty
               End If
               
           ' For DWORDS
           Case REG_DWORD:
               lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
               If lrc = ERROR_NONE Then vValue = lValue
               
           Case Else
               'other data types not supported
               lrc = -1
       End Select

QueryValueExExit:
       QueryValueEx = lrc
       Exit Function

QueryValueExError:
       Resume QueryValueExExit
   End Function








