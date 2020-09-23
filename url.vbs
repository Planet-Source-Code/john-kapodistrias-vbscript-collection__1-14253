'Example script with the "Run" command
'Created by John Kapodistrias 
'of JNKSTUDIOS
dim X
dim A
on error resume next
A=inputbox ("Enter the URl to visit","URL")
set X=createobject("Wscript.Shell")
X.Run A


