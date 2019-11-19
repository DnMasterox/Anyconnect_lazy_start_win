Function start_vpn()
  Set WshShell = Wscript.CreateObject("WScript.Shell")
  WshShell.run """%PROGRAMFILES(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpncli.exe"""
  WScript.Sleep 5000
		WshShell.SendKeys "server"
  WshShell.SendKeys "{ENTER}"
  WshShell.SendKeys "{ENTER}"
  WshShell.SendKeys getPassword()
  WScript.Sleep 5000
  WshShell.SendKeys "{ENTER}"
End Function

Function stop_vpn()
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "taskkill /f /im vpnui.exe /t", , True
End Function

Function getPassword()
	Const ForReading = 1
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.OpenTextFile("C:\creds.txt", ForReading)
    getPassword = objTextFile.Readline
End Function

Set objWMIService = GetObject ("winmgmts:")
Set proc = objWMIService.ExecQuery("select * from Win32_Process Where Name='vpnui.exe'")
If proc.count > 0 Then 
	stop_vpn()
Else 
    start_vpn()
End If
