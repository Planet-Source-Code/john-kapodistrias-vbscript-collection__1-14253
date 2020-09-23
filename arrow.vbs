Dim W
Dim S
Dim Z

Set W = WScript.CreateObject("WScript.Shell")
S = msgbox("This script will remove the shortcut arrows",vbOkCancel,"Proceed?")
if S = vbOk then
W.RegDelete "HKCR\lnkfile\IsShortcut"
W.RegDelete "HKCR\piffile\IsShortcut"
msgbox "Done"
else wscript.Quit
end if

