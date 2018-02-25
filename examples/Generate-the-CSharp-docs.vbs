With CreateObject("VBScripting.Includer")
    Execute .Read("DocGeneratorCS")
End With
With New DocGeneratorCS
    .OutputFile = "..\docs\CSharpClasses"
    .XmlFolder = "..\.Net\lib"
    .Generate
    .ViewHtml
''''.ViewMarkdown
End With
