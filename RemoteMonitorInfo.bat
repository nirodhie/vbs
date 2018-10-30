rem psexec \\ltpolbn7340 wscript.exe -nologo -c Get_Monitor_Info.vbs
net use a: /delete /yes
net use a: \\ltpolbn7340\c$
copy Get_Monitor_Info.vbs a:\temp
psexec \\ltpolbn7340 cscript.exe //nologo c:\temp\Get_Monitor_Info.vbs
pause