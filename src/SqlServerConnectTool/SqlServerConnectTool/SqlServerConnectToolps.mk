
SqlServerConnectToolps.dll: dlldata.obj SqlServerConnectTool_p.obj SqlServerConnectTool_i.obj
	link /dll /out:SqlServerConnectToolps.dll /def:SqlServerConnectToolps.def /entry:DllMain dlldata.obj SqlServerConnectTool_p.obj SqlServerConnectTool_i.obj \
		kernel32.lib rpcndr.lib rpcns4.lib rpcrt4.lib oleaut32.lib uuid.lib \

.c.obj:
	cl /c /Ox /DWIN32 /D_WIN32_WINNT=0x0400 /DREGISTER_PROXY_DLL \
		$<

clean:
	@del SqlServerConnectToolps.dll
	@del SqlServerConnectToolps.lib
	@del SqlServerConnectToolps.exp
	@del dlldata.obj
	@del SqlServerConnectTool_p.obj
	@del SqlServerConnectTool_i.obj
