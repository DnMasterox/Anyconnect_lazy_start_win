Function start_vpn()
  Dim credentials
  credentials = getDataFromTxt("credentials.txt")
  Set WshShell = Wscript.CreateObject("WScript.Shell")
  WshShell.run """%PROGRAMFILES(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpncli.exe"""
  WScript.Sleep 5000
  For i=0 to UBound(credentials)
        WshShell.SendKeys credentials(i)
        WshShell.SendKeys "{ENTER}"
        WScript.Sleep 2000
  Next
  WScript.Sleep 5000
  WshShell.SendKeys "{ENTER}"
End Function

Function stop_vpn()
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "taskkill /f /im vpnui.exe /t", , True
End Function

Function getDataFromTxt(ByVal var)
	Const ForReading = 1
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.OpenTextFile("C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\" & var, ForReading)
    Dim arrFileName
    FileContent = objTextFile.ReadAll
    arrFileName = Split(FileContent, VbCrLF)
    objTextFile.Close
    Set objOpen = Nothing
    Set objFSO = Nothing
    getDataFromTxt = arrFileName
End Function

Set objWMIService = GetObject ("winmgmts:")
Set proc = objWMIService.ExecQuery("select * from Win32_Process Where Name='vpnui.exe'")
If proc.count > 0 Then 
	stop_vpn()
Else
   start_vpn()
End If

