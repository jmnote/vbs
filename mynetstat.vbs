set fso = createobject("scripting.filesystemobject") 
set sh  = createobject("wscript.shell") 
set oExec = sh.exec ("tasklist") 
do while not oExec.stdout.atendofstream 
	line = oExec.stdout.readline 
	processes = processes & trim(left(line,27)) & "," 
	PIDs = PIDs & trim(mid(line,30,6)) & "," 
loop
PID = split(PIDs,",") 
Process = split(processes,",") 
set oExec = sh.exec ("netstat -ano") 
wscript.echo "Port"&vbtab&"PID"&vbtab&"CMD"
do while not oExec.stdout.atendofstream 
	line = oExec.stdout.readline 
	flag = true
	if mid(line,3,3)<>"TCP" then flag=false
	if mid(line,10,1)<>"0" then flag=false
	if mid(line,56,9)<>"LISTENING" then flag=false
	if flag = true then
		port = trim(mid(line,18,5))
		netpid = trim(mid(line,72))
		for x = 0 to ubound(PID)
	  	if netpid = PID(x) then exit for 
		next
		cmd = "-"
		if x <= ubound(PID) then cmd = Process(x)
		wscript.echo port&vbtab&netpid&vbtab&cmd
	end if
loop 
