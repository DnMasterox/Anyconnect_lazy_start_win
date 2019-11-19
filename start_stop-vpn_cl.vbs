Dim strAnyconnectPath
Dim strCredentialsFileName
Dim credentials

strAnyconnectPath = "C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\"
strCredentialsFileName = "credentials.txt"
credentials = getDataFromTxt(strCredentialsFileName)


Function start_vpn(ByVal creds)
  Set WshShell = Wscript.CreateObject("WScript.Shell")
  WshShell.run """%PROGRAMFILES(x86)%\Cisco\Cisco AnyConnect Secure Mobility Client\vpncli.exe"""
  WScript.Sleep 5000
  For i=0 to UBound(creds)
        WshShell.SendKeys creds(i)
        WshShell.SendKeys "{ENTER}"
        WScript.Sleep 2000
  Next
End Function

Function stop_ui_vpn()
	Set objShell = WScript.CreateObject("WScript.Shell")
	objShell.Run "taskkill /f /im vpnui.exe /t", , True
End Function

Function getDataFromTxt(ByVal var)
	Const ForReading = 1
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.OpenTextFile(strAnyconnectPath & var, ForReading)
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
	stop_ui_vpn()
Else
   start_vpn(credentials)
End If

