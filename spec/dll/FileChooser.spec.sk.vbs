
'test FileChooser.dll

Option Explicit : Initialize

With New TestingFramework

    .describe "FileChooser.dll (uses SendKeys)"

    .ShowSendKeysWarning

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
        WScript.Sleep 1000
        'prepare the input text stream
        Set input_ = fso.OpenTextFile(txtFixture, ForReading)
        .AssertEqual input_.ReadLine, anyFile

    .it "should return multiple filespecs"
        caption = "Choose two files"
        fc.Title = caption
        txtFixture = fixtureBase & "s.txt"
        vbsFixture = fixtureBase & "s.vbs"
        anyFile = WScript.ScriptFullName
        anotherFile = fso.GetAbsolutePathName(vbsFixture)
        sh.Run vbsFixture
        .MessageAppeared caption, 5, "%n%n""" & anyFile & """ """ & anotherFile & """{Enter}"
        WScript.Sleep 1000
        input_.Close
        Set input_ = fso.OpenTextFile(txtFixture, ForReading)
        files = Split(input_.ReadAll, vbCrLf)
        .AssertEqual files(0) & files(1), anyFile & anotherFile

    .CloseSendKeysWarning
End With

Cleanup

Sub Cleanup
    input_.Close
    Dim files : files = Array("" _
        , fixtureBase & ".txt" _
        , fixtureBase & "s.txt" _
    )
    Dim i
    For i = 1 To UBound(files)
        If fso.FileExists(files(i)) Then fso.DeleteFile(files(i))
    Next
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
    With CreateObject("VBScripting.Includer")
        ExecuteGlobal .read("TestingFramework")
    End With
End Sub
        

