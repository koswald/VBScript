
'test FileChooser.dll

Option Explicit : Initialize

With New TestingFramework

    .describe "FileChooser.dll"

    .it "should have a Title property"
        On Error Resume Next
            x = fc.Title
            fc.Title = x
            .AssertEqual Err.Description, ""
    .it "should have a Multiselect property"
        Err.Clear
            x = fc.Multiselect
            fc.Multiselect = x
            .AssertEqual Err.Description, ""
    .it "should have a DereferenceLinks property"
        Err.Clear
            x = fc.DereferenceLinks
            fc.DereferenceLinks = x
            .AssertEqual Err.Description, ""
    .it "should have a DefaultExt property"
        Err.Clear
            x = fc.DefaultExt
            fc.DefaultExt = x
            .AssertEqual Err.Description, ""
    .it "should have a Filter property"
        Err.Clear
            x = fc.Filter
            fc.Filter = x
            .AssertEqual Err.Description, ""
    .it "should have a FilterIndex property"
        Err.Clear
            x = fc.FilterIndex
            fc.FilterIndex = x
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .it "should open a chooser dialog"
        caption = "Choose a single file"
        fc.Title = caption
        fixtureBase = "fixture\FileChooserGetFile"
        txtFixture = fixtureBase & ".txt"
        vbsFixture = fixtureBase & ".vbs"
        anyFile = WScript.ScriptFullName
        'run the fixture file to open the dialog
        sh.Run vbsFixture
        'wait for dialog, enter file, acknowledge dialog
        .AssertEqual .MessageAppeared(caption, 5, _
            "%n%n""" & anyFile & """{Enter}"), True

    .it "should return a single filespec"
        'wait for fixture to write result to file
        'and close the text stream
        WScript.Sleep 500
        'prepare the input text stream
        Set input_ = fso.OpenTextFile(txtFixture, ForReading)
        .AssertEqual input_.ReadLine, anyFile

    .it "should return multiple filespecs"
        caption = "Choose two files"
        fc.Title = caption
        fixtureBase = "fixture\FileChooserGetFiles"
        txtFixture = fixtureBase & ".txt"
        vbsFixture = fixtureBase & ".vbs"
        anyFile = WScript.ScriptFullName
        anotherFile = fso.GetAbsolutePathName(vbsFixture)
        'run the fixture file to open the dialog
        sh.Run vbsFixture
        'wait for dialog, enter file, acknowledge dialog
        .MessageAppeared caption, 5, "%n%n""" & anyFile & """ """ & anotherFile & """{Enter}"
        'wait for fixture to write result to file
        'and close the text stream
        WScript.Sleep 500
        'prepare the input text stream
        input_.Close
        Set input_ = fso.OpenTextFile(txtFixture, ForReading)
        files = Split(input_.ReadAll, vbCrLf)
        .AssertEqual files(0) & files(1), anyFile & anotherFile

End With

Cleanup

Sub Cleanup
    input_.Close
    Set input_ = Nothing
    Set fc = Nothing
    Set sh = Nothing
    Set fso = Nothing
End Sub

Dim fc, fso, sh
Dim x
Dim caption
Dim fixture, input_, inFile
Dim fixtureBase, txtFixture, vbsFixture
Dim anyFile, anotherFile, files
Const ForReading = 1

Sub Initialize
    Set fc = CreateObject("VBScripting.FileChooser")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sh = CreateObject("WScript.Shell")
    With CreateObject("includer")
        ExecuteGlobal .read("TestingFramework")
    End With
End Sub
        

