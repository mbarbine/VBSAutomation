Option Explicit

'Delcare Variables

Dim strCorpProxyServer '
Dim strProxyOverRide ' 
Dim strSuffixSearchList ' String for Suffix Search List '
Dim objNetwork ' Network Object
Dim objWSHShell ' Object into Windows Shell

'ERROR HANDLING' 

on error resume next
Dim SysVarReg, Value, Value2
Value = ""
Value2 = "2"
Set SysVarReg = WScript.CreateObject("WScript.Shell")
SysVarReg.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\", Value, "REG_SZ" 
SysVarReg.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\HTTP", Value2, "REG_DWORD" 

'Setting Registry changes information

strCorpProxyServer = ""
strProxyOverRide = ""
strSuffixSearchList = ""

Set objWSHShell = WScript.CreateObject ("WScript.Shell")

'Setting the Proxy Server, Proxy OverRideList, and the Suffix Search Order list via the Registry

objWSHShell.RegWrite ("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer"), strCorpProxyServer, "REG_SZ"
objWSHShell.RegWrite ("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable"), "0", "REG_DWORD"
objWSHShell.RegWrite ("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyOverride"), strProxyOverRide, "REG_SZ"

'Suffix Search Order List
objWSHShell.RegWrite "HKLM\System\CurrentControlSet\Services\TCPIP\Parameters\SearchList", strSuffixSearchList, "REG_SZ"


msgbox("Done! Proxy is Disabled!") 
