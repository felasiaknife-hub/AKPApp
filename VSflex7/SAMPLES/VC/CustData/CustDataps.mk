
CustDataps.dll: dlldata.obj CustData_p.obj CustData_i.obj
	link /dll /out:CustDataps.dll /def:CustDataps.def /entry:DllMain dlldata.obj CustData_p.obj CustData_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del CustDataps.dll
	@del CustDataps.lib
	@del CustDataps.exp
	@del dlldata.obj
	@del CustData_p.obj
	@del CustData_i.obj
