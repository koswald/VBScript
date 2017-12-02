
'fixture for FileChooser.spec.vbs

'open file chooser dialog, write results to file

Set fc = CreateObject("VBScripting.FileChooser")
fc.Title = "Choose two files"
fc.Multiselect = True
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.OpenTextFile("fixture\FileChooserGetFiles.txt", ForWriting, CreateNew)
files = fc.FileNames
For Each file In files
    out.WriteLine file
Next
out.Close
Set out = Nothing
Set fso = Nothing
Set fc = Nothing
Const ForWriting = 2
Const CreateNew = True
