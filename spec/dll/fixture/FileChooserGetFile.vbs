
'fixture for FileChooser.spec.vbs

'open file chooser dialog, and write result/response to file

Set fc = CreateObject("VBScripting.FileChooser")
fc.Title = "Choose a single file"
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.OpenTextFile("fixture\FileChooserGetFile.txt", ForWriting, CreateNew)
out.WriteLine fc.FileName
out.Close
Set out = Nothing
Set fso = Nothing
Set fc = Nothing
Const ForWriting = 2
Const CreateNew = True
