With CreateObject("VBScripting.Includer")
    Execute .Read("DocGeneratorCS")
End With
Set fso = CreateObject("Scripting.FileSystemObject")
With CreateObject("WScript.Shell")
    .CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
End With
Set fso = Nothing
With New DocGeneratorCS
    .OutputFile = "..\docs\CSharpClasses"
    .XmlFolder = "..\.Net\lib"
    .Generate
'    .ViewHtml
End With
