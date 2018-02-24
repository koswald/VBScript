With CreateObject("VBScripting.Includer")
    Execute .Read("DocGeneratorCS")
End With
With New DocGeneratorCS
    .HtmlFile = "..\docs\CSharpClasses.html"
    .XmlFolder = "..\.Net\lib"
    .Generate
    .View
End With
