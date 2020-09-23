'Script example with the CommonDialog control with the run command
'Created by John Kapodistrias 
'of JNKSTUDIOS

Option Explicit

Dim Con
Dim S 
Dim Ret
Dim X
Dim Q

Sub CDLG 
Set Con=Wscript.CreateObject("MSComDLg.CommonDialog") 'create an instance of the control
With con 'set the parameters
	.DialogTitle="VBS Common Dialog Example"
	.MaxFileSize=260
	.Filter="All files|*.*"
	.ShowOpen 
End With
s=con.FileName
 if s = "" then   'do a check if a file has been selected
   msgbox "No file selected"
   wscript.Quit
 end if
End Sub

Call CDLG  'call the procedure

msgbox s

ret = msgbox("Would you like to run the selected file?",vbOkCancel,"Run?") 
 if ret = vbOkCancel then  
 Set X=Wscript.CreateObject("Wscript.Shell") 'create an instance of the shell object
 x.Run s  'run the file
end if

'End of Script