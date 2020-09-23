Dim Path
Dim Con
Dim X
Dim R

Sub CDLG 
Set Con=Wscript.CreateObject("MSComDLg.CommonDialog") 
With con 
	.DialogTitle="Select a bitmap"
	.MaxFileSize=260
	.Filter="Bitmap files|*.bmp"
	.ShowOpen 
End With
Path = con.FileName
 if Path = "" then   
   msgbox "No file selected"
   wscript.Quit
 end if
End Sub

X = msgbox("You are about to change the background of Explorer.Proceed?",vbOkCancel,"Confirm")
if x = vbOk then
Call CDLG
Set R = CreateObject("Wscript.Shell")
r.RegWrite "HKCU\SOFTWARE\Microsoft\Internet Explorer\Toolbar\BackBitmap",Path
Call msgbox("Done.",vbOkOnly,"Done")
else
wscript.Quit
end if


