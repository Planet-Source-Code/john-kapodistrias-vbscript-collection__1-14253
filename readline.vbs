Dim Con
Dim Ret
Dim S
Sub CDLG 
Set Con=Wscript.CreateObject("MSComDLg.CommonDialog") 'create an instance of the control
With con 'set the parameters
	.DialogTitle="Select a text file"
	.MaxFileSize=260
	.Filter="Text files|*.txt"
	.ShowOpen 
End With
S = con.FileName
 if s = "" then   'do a check if a file has been selected
   msgbox "No file selected"
   wscript.Quit
 end if
End Sub

Function ReadAllTextFile
   Const ForReading = 1, ForWriting = 2
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile(S, ForReading)
   ReadAllTextFile = f.ReadAll
    if S = "" then 
   Msgbox "The text file selected is empty"
   Wscript.Quit
   End If
End Function

Call CDLG

Msgbox ReadAllTextFile