
CellChangeps.dll: dlldata.obj CellChange_p.obj CellChange_i.obj
	link /dll /out:CellChangeps.dll /def:CellChangeps.def /entry:DllMain dlldata.obj CellChange_p.obj CellChange_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del CellChangeps.dll
	@del CellChangeps.lib
	@del CellChangeps.exp
	@del dlldata.obj
	@del CellChange_p.obj
	@del CellChange_i.obj
