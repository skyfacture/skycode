#l::

msg =
(
You pressed Win+L.

The Win+L lock key is disabled via the registry setting:

reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System" /v DisableLockWorkstation /t REG_DWORD /d 1 /f

To re-enable locking:

reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\System" /v DisableLockWorkstation /f
)

MsgBox, %msg%
return
