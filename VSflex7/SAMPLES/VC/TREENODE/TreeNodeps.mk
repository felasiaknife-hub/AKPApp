
TreeNodeps.dll: dlldata.obj TreeNode_p.obj TreeNode_i.obj
	link /dll /out:TreeNodeps.dll /def:TreeNodeps.def /entry:DllMain dlldata.obj TreeNode_p.obj TreeNode_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del TreeNodeps.dll
	@del TreeNodeps.lib
	@del TreeNodeps.exp
	@del dlldata.obj
	@del TreeNode_p.obj
	@del TreeNode_i.obj
