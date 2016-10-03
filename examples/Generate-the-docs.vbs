
With CreateObject("includer")
	Execute(.read("DocGenerator"))
End With

Set gen = New DocGenerator

gen.SetTitle "Karl's VBScript utilities"
gen.SetScriptFolder "../class" 'folders are set relative to this script file's location
gen.SetDocFolder ".."
gen.SetFilesToDocument(".*\.(vbs|wsf|wsc)")
gen.SetDocName "TheDocs.html"
gen.Generate
gen.View
