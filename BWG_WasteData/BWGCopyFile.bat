CD\
CLS

@Echo Makking BWG Directory in C:\Windows\Temp\
MKDIR C:\Windows\Temp\BWG

Copy \\192.168.2.2\AppUpdate\Upload\*.*  C:\Windows\Temp\BWG\*.* /Y

Exit

