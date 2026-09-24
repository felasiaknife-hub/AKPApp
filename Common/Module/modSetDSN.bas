Attribute VB_Name = "modSetDSN"
Option Explicit

Public Function SetDSN(ByVal DB_Name As String, _
                        ByVal DSN As String, _
                        ByVal Description As String, _
                        ByVal Driver_Name As String, _
                        ByVal Driver_Path As String, _
                        ByVal Last_User As String, _
                        ByVal Server_Name As String, _
                        ByRef Status As String _
                        ) As Boolean

   Dim ThisODBC As New CLSODBC
   Dim Msg As String
   
    ThisODBC.DatabaseName = DB_Name
    ThisODBC.DataSourceName = DSN
    
    ThisODBC.Description = Description
    
    ThisODBC.DriverName = Driver_Name
    
    ThisODBC.DriverPath = Driver_Path
    ThisODBC.LastUser = Last_User
    ThisODBC.USERID = Last_User
    ThisODBC.Server = Server_Name
    ThisODBC.Servername = Server_Name
    SetDSN = ThisODBC.SetDSN
    'SetDSN = ThisODBC.SetUserDSN
    
    Status = ThisODBC.Status
    
    Set ThisODBC = Nothing
    
End Function

