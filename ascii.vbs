Function ASCIIConvert(SourceString)
On Error Resume Next
    Dim CurChr
    ASCIIConvert = ""
    For i = 1 To Len(SourceString)
        CurChr = Mid(SourceString, i, i + 1)
        ASCIIConvert = ASCIIConvert & " " & Asc(CurChr)
    Next 
    ASCIIConvert = Right(ASCIIConvert, Len(ASCIIConvert) - 3)
End Function

dim X 
dim S
x=inputbox ("Enter your text","Text")
S = asciiconvert(x)
msgbox S
