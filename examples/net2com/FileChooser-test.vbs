
'test that FileChooser.cs compiled correctly and
'that FileChooser.dll registered correctly

'expected outcome:
' A browse for a file dialog appears.
' In the dropdown box above the Open button,
' "All files (*.*)" is selected automatically.
' The list item above it should be "RTF files".
' After a file is manually selected, and Open is clicked,
' a message box displays the filespec of the selected file.

' After OK is clicked, browse for a file dialog appears again.
' After two or more files are manually selected, and Open is clicked,
' a message box displays the filespecs of the selected files.

Dim fc : Set fc = CreateObject("FileChooser")

fc.Filter = "RTF files | *.rtf | All files | *.*"
fc.FilterIndex = 2

MsgBox "Title: "  & fc.Title & vbLf & _
        "Multiselect: " & fc.Multiselect & vbLf & _
        "DereferenceLinks: " & fc.DereferenceLinks & vbLf & _
        "DefaultExt: " & fc.DefaultExt

fc.Title = "Testing: Select a single file."
MsgBox "Single file: " & vbLf & fc.FileName, _
                vbInformation, WScript.ScriptName

fc.Title = "Testing: Select multiple files."
fc.Multiselect = True
Dim files, i, s : s = "" : files = fc.FileNames
For i = 0 To UBound(files)
    s = s & vbLf & files(i)
Next
MsgBox "Multi files:" & s, vbInformation, _
                WScript.ScriptName

Set fc = Nothing
