
'generate a documentation file based on code comments

Set includer = CreateObject("VBScripting.Includer")
Execute includer.read("DocGenerator")
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
Set fso = Nothing

With New DocGenerator
 .SetTitle "VBScript Utility Classes Documentation"
 .SetScriptFolder "..\class" 'location of the scripts to document, relative to this script
 .SetFilesToDocument "*.vbs | *.wsf | *.wsc" 'filename(s) of the scripts to document
 .SetDocFolder "..\docs" 'location of the target documentation file, relative to this script
 .SetDocName "VBScriptClasses.html"
 .Generate
End With

sh.PopUp "Done generating the VBScripting classes docs.", 3, WScript.ScriptName, vbInformation + vbSystemModal
Set includer = Nothing
Set sh = Nothing

