' Usage: > cscript//nologo product.vbs
set sh  = createobject("wscript.shell") 
set oExec = sh.exec ("wmic product") 
wscript.echo "Product"
do while not oExec.stdout.atendofstream 
	line = oExec.stdout.readline 
	line = trim(mid(line, InStr(line, "  ")+1, Len(line)))
	line = trim(left(line, InStr(line, "  ")))
	wscript.echo line
loop
