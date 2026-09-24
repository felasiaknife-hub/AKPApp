
VSPPG7ps.dll: dlldata.obj VSPPG7_p.obj VSPPG7_i.obj
	link /dll /out:VSPPG7ps.dll /def:VSPPG7ps.def /entry:DllMain dlldata.obj VSPPG7_p.obj VSPPG7_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del VSPPG7ps.dll
	@del VSPPG7ps.lib
	@del VSPPG7ps.exp
	@del dlldata.obj
	@del VSPPG7_p.obj
	@del VSPPG7_i.obj
