Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
current = fso.GetAbsolutePathName(".")
sh.Run """" & current & "\Start2.bat" & """", 0, False
Set sh = Nothing
Set fso = Nothing
