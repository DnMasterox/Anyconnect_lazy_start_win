Function start_vpn()
  Set WshShell = Wscript.CreateObject("WScript.Shell")
  WshShell.run """%PROGRAMFILES(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpnui.exe"""
  WScript.Sleep 5000
  WshShell.AppActivate "Cisco AnyConnect Secure Mobility Client"
  WshShell.SendKeys "{ENTER}"
  WScript.Sleep 8000
  WshShell.SendKeys "{!}{%}{[]}"
  WshShell.SendKeys "{ENTER}"
End Function

Function stop_vpn()
	objShell.Run "taskkill /f /im vpnui.exe /t", , True
End Function

Set objShell = WScript.CreateObject("WScript.Shell")
Set objWMIService = GetObject ("winmgmts:")
Set proc = objWMIService.ExecQuery("select * from Win32_Process Where Name='vpnui.exe'")
If proc.count > 0 Then 
	stop_vpn()
Else 
start_vpn()
End If