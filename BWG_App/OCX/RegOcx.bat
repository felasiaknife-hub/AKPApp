@REM FOR /f %%a IN ('dir /b *.ocx') DO  (@echo Registering file %%a & xcopy /y %%a %SystemDirectory%\*.*)
@FOR /f %%a IN ('dir /b *.ocx') DO  (
@rem echo Copying file %%a to %SystemRoot%\System32\%%a
copy /y %%a %SystemRoot%\System32\%%a 
@rem echo Registering file %%a 
Regsvr32 /s %%a
)
@PAUSE
