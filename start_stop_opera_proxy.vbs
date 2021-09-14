Option Explicit
Dim valUserIn
Dim objShell, RegLocate, RegLocate1
Set objShell = WScript.CreateObject("WScript.Shell")
On Error Resume Next
valUserIn = MsgBox("Use Opera Proxy?",4,"Opera Select")
If valUserIn=vbYes Then
objShell.Run "opera-proxy.windows-amd64.exe"
RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer"
objShell.RegWrite RegLocate,"127.0.0.1:18080","REG_SZ"
RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable"
objShell.RegWrite RegLocate,"1","REG_DWORD"
MsgBox "Opera Proxy is Enabled"
else
RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer"
objShell.RegWrite RegLocate,"0.0.0.0:80","REG_SZ"
RegLocate = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable"
objShell.RegWrite RegLocate,"0","REG_DWORD"
MsgBox "Opera Proxy is Disabled"
End If
WScript.Quit