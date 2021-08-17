Dim WMIService, Process, Processes, Flag, WS
Set WMIService = GetObject("winmgmts:{impersonationlevel=impersonate}!\\.\root\cimv2")
Set WS = Wscript.CreateObject("Wscript.Shell")
trial = 0
Do while trial < 20 '随便设的，防止进入死循环
	count = 0 '正在运行的rclone进程数量
	trial = trial + 1
	Set Processes = WMIService.ExecQuery("select * from win32_process")
	for each Process in Processes
		if strcomp(Process.name, "rclone.exe") = 0 then
			count = count + 1
		end if
	next
	If count = 0 Then
		'工作盘，缓存大，读写快
		WS.Run "rclone mount --volname Personal sftp://personal D:    --cache-dir %temp%  --vfs-cache-mode full --attr-timeout 10m --vfs-cache-max-age 24h --vfs-cache-max-size 1G --buffer-size 200M", 0
	ElseIf count = 1 Then
		'资源盘，只读防止误删
		WS.Run "rclone mount --volname Cloud  sftp://share E:   --read-only --cache-dir %temp%  --vfs-cache-mode writes", 0
	Else Exit Do
	End If
	WS.Sleep 500 '等待挂载完成
Loop
Set WMIService = nothing
