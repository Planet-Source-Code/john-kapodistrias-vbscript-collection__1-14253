Function GetScriptEngineInfo
   Dim s
   s = ""   ' Build string with necessary info.
   s = ScriptEngine & " Version "
   s = s & ScriptEngineMajorVersion & "."
   s = s & ScriptEngineMinorVersion & "."
   s = s & ScriptEngineBuildVersion 
   GetScriptEngineInfo = s   ' Return the results.
End Function
call GetScriptEngineInfo
msgbox GetScriptEngineInfo