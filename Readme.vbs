
Function ReadAllTextFile
   Const Read = 1, Write = 2
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile("readme.txt", Read)
   ReadAllTextFile = f.ReadAll
End Function

Call msgbox (ReadAllTextFile,vbOkOnly,"Contents")
