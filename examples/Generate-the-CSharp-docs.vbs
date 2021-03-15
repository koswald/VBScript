Set includer = CreateObject("VBScripting.Includer")
Execute includer.Read("DocGeneratorCS")
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
sh.CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
Set fso = Nothing

With New DocGeneratorCS
    .OutputFile = "..\docs\CSharpClasses"
    .XmlFolder = "..\.NET\lib"
    .Generate
End With

sh.PopUp "Done generating the C# classes docs.", 3, WScript.ScriptName, vbInformation + vbSystemModal
Set sh = Nothing
Set includer = Nothing
