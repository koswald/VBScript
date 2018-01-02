
'test VBScripting.FolderChooser
' and VBScripting.FolderChooser2

Option Explicit : Initialize

'troubleshooting intermittent spec failure,
'in which dialog opens but does not close...
'"fix": change exitPause from 0 to 500

Const exitPause = 500 'milliseconds

Main
Cleanup

Sub Main
    Dim i : For i = 0 To UBound(progIds)
        TestByProgIdIndex(i)
    Next
End Sub

Sub TestByProgIdIndex(index)
    Dim progId : progId = progIds(index)
    Dim caption 'title-bar text
    Const defaultCaption = "Select a folder"
    Dim o 'object under test
    Dim tsIn 'text stream for reading
    Dim actual, expected
    Dim keys 'keystrokes to send to the dialog

    With New TestingFramework
        .ShowSendKeysWarning

        .describe progId
            Set o = CreateObject(progId)

        .it "should default to open at the current directory"
            .AssertEqual sh.CurrentDirectory, o.InitialDirectory

        .it "should support environment variables"
            sh.Run format(Array( _
                 "%s.vbs %s %UserProfile% .", _
                fixtureBase, progId _
            ))
            keys = "{tab}{tab}{tab}{tab}{tab}{tab}{tab}{ }{tab}{tab}{tab}{tab}{enter}" 'TabX7 to left pane, select the initial folder (space), TabX4 to "Select folder" button, click the button (enter)
            If .MessageAppeared(defaultCaption, 5, keys) Then
                Set tsIn = GetReadStream(fixtureBase & ".txt")
                actual = tsIn.ReadLine
                expected = sh.ExpandEnvironmentStrings("%UserProfile%")
                .AssertEqual actual, expected
                tsIn.Close
                Set tsIn = Nothing
            Else Err.Raise 1,, "Fixture dialog """ & caption & """ failed to open"
            End If
            DeleteTempFile(fixtureBase & ".txt")

        .it "should support relative paths"
            caption = progId & " test" 'simultaneously test the Title method
            sh.Run format(Array( _
                 "%s.vbs %s .. ""%s""", _
                fixtureBase, progId, caption _
            ))
            keys = "{tab}{tab}{tab}{tab}{tab}{tab}{tab}{ }{tab}{tab}{tab}{tab}{enter}" 'TabX7 to left pane, select the initial folder (space), TabX4 to "Select folder" button, click the button (enter)
            If .MessageAppeared(caption, 5, keys) Then
                Set tsIn = GetReadStream(fixtureBase & ".txt")
                actual = tsIn.ReadLine
                expected = GetParent(GetParent(WScript.ScriptFullName))
                .AssertEqual actual, expected
                tsIn.Close
                Set tsIn = Nothing
            Else Err.Raise 2,, "Fixture dialog """ & caption & """ failed to open"
            End If
            DeleteTempFile(fixtureBase & ".txt")

            .CloseSendKeysWarning
    End With
    Set o = Nothing
End Sub

Function GetParent(item)
    GetParent = fso.GetParentFolderName(item)
End Function

Sub ReleaseObjectMemory
    Set sh = Nothing
    Set fso = Nothing
End Sub

Sub Cleanup
    ReleaseObjectMemory
End Sub

Sub DeleteTempFile(tmpFile)
    If fso.FileExists(tmpFile) Then fso.DeleteFile(tmpFile)
End Sub

'Wait for the fixture to finish writing to the output file,
'then return a text stream for reading the output file.
Function GetReadStream(file)
    Const loopPause = 25 'milliseconds
    Const nLoops = 200 'max number of loops
    Dim i : i = 0
    On Error Resume Next
        Do
            Err.Clear
            Set GetReadStream = fso.OpenTextFile(file, ForReading)
            If Err = 0 Then
                'the fixture is done writing because there was 
                'no error attempting to open the file for reading
                WScript.Sleep exitPause
                Exit Function
            End If
            WScript.Sleep loopPause
            i = i + 1
        Loop Until i = nLoops
    On Error Goto 0
    Err.Raise 3,, "Couldn't open file """ & file & """"
End Function

Const idStrings = "VBScripting.FolderChooser|VBScripting.FolderChooser2"
Const fixtureBase = "fixture\FolderChooserX.fixture"
Const ForReading = 1
Dim progIds
Dim format
Dim sh, fso

Sub Initialize
    progIds = Split(idStrings, "|")
    With CreateObject("includer")
        Execute .read("StringFormatter")
        ExecuteGlobal .read("TestingFramework")
    End With
    Set format = New StringFormatter
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub
