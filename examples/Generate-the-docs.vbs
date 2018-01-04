
'generate a documentation file based on code comments

With CreateObject("includer")
    Execute .read("DocGenerator")
End With

With New DocGenerator
    .SetTitle "VBScript Utility Classes Documentation"
    .SetScriptFolder "..\..\class" 'location of the scripts to document, relative to this script
    .SetFilesToDocument "*.vbs | *.wsf | *.wsc" 'filename(s) of the scripts to document
    .SetDocFolder "..\..\docs" 'location of the target documentation file, relative to this script
    .SetDocName "TheDocs.html"
    .Generate
    .View
End With
