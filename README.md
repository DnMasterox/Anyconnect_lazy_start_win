"lazy connection scripts"
=====================
Lazy connection scripts for Windows 

Note:
-----------------------------------
You need to create "credentials.txt" file at 
`C:\Program Files (x86)\Cisco\Cisco AnyConnect Secure Mobility Client\`
or define path in ***strAnyconnectPath*** variable and put the name of server and your credentials 
to this file in order:
<li> connect {server name}
<li> {user name} 
<li> {password}

Do not forget to shield symbols for password with "{ }"

Run "Cisco Anyconnect" from app
---------
Lazy one click start/stop script for Windows
Just Execute the script 'start_stop-vpn.vbs' 
Do not touch mouse and keyboard

Run from command line
---------
Lazy one click start script for Windows
Execute the script. 'start_stop-vpn_cl.vbs' 

Warning
---------
It is not secured to store your password in *.vbs or *.txt files.
"Stop" function is not safe and could make damage