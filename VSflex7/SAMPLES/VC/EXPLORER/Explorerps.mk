
Explorerps.dll: dlldata.obj Explorer_p.obj Explorer_i.obj
	link /dll /out:Explorerps.dll /def:Explorerps.def /entry:DllMain dlldata.obj Explorer_p.obj Explorer_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del Explorerps.dll
	@del Explorerps.lib
	@del Explorerps.exp
	@del dlldata.obj
	@del Explorer_p.obj
	@del Explorer_i.obj
