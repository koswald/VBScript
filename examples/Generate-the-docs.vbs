
With CreateObject("includer")
    Execute(.read("DocGenerator"))
End With

With New DocGenerator
    .SetTitle "Karl's VBScript utilities"
    .SetScriptFolder "../class" 'relative to this script
    .SetDocFolder ".."
    .SetFilesToDocument(".*\.(vbs|wsf|wsc)")
    .SetDocName "TheDocs.html"
    .Generate
    .View
End With
