Dim WMIService, Process, Processes, Flag, WS
Set WMIService = GetObject("winmgmts:{impersonationlevel=impersonate}!\\.\root\cimv2")
Set WS = Wscript.CreateObject("Wscript.Shell")
trial = 0
Do while trial < 20
	count = 0
	trial = trial + 1
	Set Processes = WMIService.ExecQuery("select * from win32_process")
	for each Process in Processes
		if strcomp(Process.name, "rclone.exe") = 0 then
			count = count + 1
		end if
	next
	If count = 0 Then
		WS.Run "rclone mount --volname Personal sftp://srv/725d030c-773b-4b5c-a7ee-a99d8f8a3191/switch/LiMC D:    --cache-dir %temp%  --vfs-cache-mode full --attr-timeout 10m --vfs-cache-max-age 24h --vfs-cache-max-size 1G --buffer-size 200M", 0
	ElseIf count = 1 Then
		WS.Run "rclone mount --volname Cloud  sftp://srv/725d030c-773b-4b5c-a7ee-a99d8f8a3191 E:   --read-only --cache-dir %temp%  --vfs-cache-mode writes", 0
	Else Exit Do
	End If
Loop
Set WMIService = nothing
