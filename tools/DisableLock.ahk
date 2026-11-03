#l::
MsgBox You pressed Win+L - Win + L lock key is disabled via registry group policy (below):

reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System" /v DisableLockWorkstation /t REG_DWORD /d 1 /f

To re-enable, use the following in powershell:

reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System" /v DisableLockWorkstation /f
return