
Selectionps.dll: dlldata.obj Selection_p.obj Selection_i.obj
	link /dll /out:Selectionps.dll /def:Selectionps.def /entry:DllMain dlldata.obj Selection_p.obj Selection_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del Selectionps.dll
	@del Selectionps.lib
	@del Selectionps.exp
	@del dlldata.obj
	@del Selection_p.obj
	@del Selection_i.obj
