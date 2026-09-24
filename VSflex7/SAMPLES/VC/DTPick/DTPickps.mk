
DTPickps.dll: dlldata.obj DTPick_p.obj DTPick_i.obj
	link /dll /out:DTPickps.dll /def:DTPickps.def /entry:DllMain dlldata.obj DTPick_p.obj DTPick_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del DTPickps.dll
	@del DTPickps.lib
	@del DTPickps.exp
	@del dlldata.obj
	@del DTPick_p.obj
	@del DTPick_i.obj
