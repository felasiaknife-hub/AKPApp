CD\
C:
CLS

@Echo Checking BWG Directory in C:\Windows\Temp\


Copy C:\WINDOWS\Temp\BWG\WasteDataForm.exe  C:\Progra~1\BWG\WasteDataForm.exe /Y
Copy C:\WINDOWS\Temp\BWG\WasteDataForm.rpt  C:\Progra~1\BWG\Reports\WasteDataForm.rpt /Y
Copy C:\WINDOWS\Temp\BWG\WasteDataFormAP.rpt  C:\Progra~1\BWG\Reports\WasteDataFormAP.rpt /Y
Copy C:\WINDOWS\Temp\BWG\dbconfig.ini  C:\Progra~1\BWG\dbconfig.ini /Y

@Echo C:\Progra~1\BWG\WasteDataForm.exe

exit