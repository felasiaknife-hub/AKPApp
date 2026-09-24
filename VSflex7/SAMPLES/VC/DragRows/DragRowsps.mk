
DragRowsps.dll: dlldata.obj DragRows_p.obj DragRows_i.obj
	link /dll /out:DragRowsps.dll /def:DragRowsps.def /entry:DllMain dlldata.obj DragRows_p.obj DragRows_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del DragRowsps.dll
	@del DragRowsps.lib
	@del DragRowsps.exp
	@del dlldata.obj
	@del DragRows_p.obj
	@del DragRows_i.obj
