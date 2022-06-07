
'test that FileChooser.cs compiled correctly and
'that FileChooser.dll registered correctly

'expected outcome:
'  see fixture\FileChooser-test-case.vbs

Option Explicit : Initialize

'access some properties to verify that they exist
s = "Title: "  & fc.Title
s = "Multiselect: " & fc.Multiselect
s = "DereferenceLinks: " & fc.DereferenceLinks
s = "DefaultExt: " & fc.DefaultExt

'general settings
fc.Filter = "RTF files | *.rtf | All files | *.*"
fc.FilterIndex = 2

'dialog settings for single file
fc.Title = "Testing: Select a single file."

'open the dialog and get the response
Dim file : file = fc.FileName

'check for user cancel
If file = "" Then Quit

'show results
msg = "Single file: " & vbLf & file
If vbCancel = MsgBox(msg, mode, caption) Then Quit

'multifile settings
fc.Title = "Testing: Select multiple files."

'open the dialog and get the response
Dim files : files = fc.FileNames

'check for user cancel
If UBound(files) = -1 Then Quit

'show results
s = ""
For i = 0 To UBound(files)
    s = s & vbLf & files(i)
Next
mode = mode - vbOKCancel + vbOKOnly
MsgBox "Multi files:" & s, mode, caption

Quit

Dim fc
Dim sh
Dim testCaseMessage
Dim mode
Dim caption
Dim msg, s, i

Sub Initialize
    Set fc = CreateObject( "VBScripting.FileChooser" )
    Set sh = CreateObject( "WScript.Shell" )
    Set testCaseMessage = _
            sh.Exec("wscript fixture\FileChooser-test-case.vbs")
    mode = vbInformation + vbSystemModal + vbOKCancel
    caption = WScript.ScriptName
End Sub

Sub Quit
    Set fc = Nothing
    testCaseMessage.Terminate
    Set testCaseMessage = Nothing
    Set sh = Nothing
    WScript.Quit
End Sub
