Dim C
Dim Con
C = msgbox ("You are about to change the CPU priority.Proceed?",vbOkCancel,"Confirm")
if C = vbOk then
Set Con = CreateObject("Wscript.Shell")
Con.RegWrite "HKLM\System\CurrentControlSet\Services\VxD\BIOS\CPUPriority",1,"REG_DWORD"
Call msgbox ("Done." & vbcrlf & "For changes to take effect you must reboot",vbOkOnly,"Done")
else Wscript.Quit
end if
