
WallPaperps.dll: dlldata.obj WallPaper_p.obj WallPaper_i.obj
	link /dll /out:WallPaperps.dll /def:WallPaperps.def /entry:DllMain dlldata.obj WallPaper_p.obj WallPaper_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del WallPaperps.dll
	@del WallPaperps.lib
	@del WallPaperps.exp
	@del dlldata.obj
	@del WallPaper_p.obj
	@del WallPaper_i.obj
