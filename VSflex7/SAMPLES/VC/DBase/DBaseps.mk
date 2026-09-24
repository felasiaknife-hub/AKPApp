
DBaseps.dll: dlldata.obj DBase_p.obj DBase_i.obj
	link /dll /out:DBaseps.dll /def:DBaseps.def /entry:DllMain dlldata.obj DBase_p.obj DBase_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del DBaseps.dll
	@del DBaseps.lib
	@del DBaseps.exp
	@del dlldata.obj
	@del DBase_p.obj
	@del DBase_i.obj
