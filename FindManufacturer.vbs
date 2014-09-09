strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colBIOS = objWMIService.ExecQuery _
("Select * from Win32_BIOS")
For each objBIOS in colBIOS
Wscript.Echo "Manufacturer: " & objBIOS.Manufacturer
Wscript.Echo "Serial Number: " & objBIOS.SerialNumber
Next

