Dim P 
Dim S
Dim X

X = msgbox ("Would you like to apply the Word 97 LiveScrolling patch?",vbOkCancel,"Confirm")
if x = vbOk
Set P = CreateObject("Wscript.Shell")
p.RegWrite "HKCU\Software\Microsoft\Office\8.0\Word\LiveScroling",1
msgbox "Done"
else wscript.Quit
