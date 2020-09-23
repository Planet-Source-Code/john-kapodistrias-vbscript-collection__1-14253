Dim Cap
Dim S
Dim X
S = msgbox("You are about to change the caption of Internet Explorer.Proceed?",vbOkCancel,"Confirm")
if S = vbOk then
X = inputbox("Enter the new caption","New caption")
Set Cap = CreateObject("Wscript.Shell")
cap.RegWrite "HKCU\SOFTWARE\Microsoft\Internet Explorer\Main\Window Title",X
else wscript.Quit
end if

