Dim m, mode

Sub FREEM()

mode = MsgBox("Do you like to free your systems physical memory?", vbYesNo, "Free Memory")
If mode = vbYes Then
m = InputBox("Enter the amount of memory in bytes", "Bytes")
FreeMem = Space(m)
Else wscript.Quit
End If
End Sub


call FREEM
