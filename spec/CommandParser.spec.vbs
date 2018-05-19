
'test the CommandParser class

With CreateObject("VBScripting.Includer")
    Execute .read("CommandParser")
    Execute .read("TestingFramework")
End With

With New TestingFramework

    .describe "CommandParser class"
        Dim cp : Set cp = New CommandParser

    .it "should raise an error if no command was specified"
        On Error Resume Next
            cp.GetResult
            .AssertErrorRaised
        On Error Goto 0

    .it "should raise an error if an invalid command was given"
        On Error Resume Next
            cp.SetCommand "this is an invalid command"
            cp.GetResult
            .AssertErrorRaised
        On Error Goto 0

    Dim x
    .it "raises an error on edge case 1a: not using cmd /c"
        On Error Resume Next
            cp.SetCommand "If defined ProgramFilesX86 (echo 64-bit) else (echo 32-bit)"
            cp.SetSearchPhrase "64-bit"
            x = cp.GetResult
            .AssertEqual "{ " & Err.Description & " }", "{ The system cannot find the file specified." & vbCrLf & " }"
        On Error Goto 0

    .it "should not raise error on edge case 1b: using cmd /c"
        On Error Resume Next
            cp.SetCommand "cmd /c If defined ProgramFilesX86 (echo 64-bit) else (echo 32-bit)"
            cp.SetSearchPhrase "64-bit"
            x = cp.GetResult
            .AssertEqual "{ " & Err.Description & " }", "{  }"
        On Error Goto 0
    
    'setup
        Dim searchPhrase1, searchPhrase2
        cp.SetCommand "xcopy /?"
        cp.SetStartPhrase "Specifies"
        cp.SetStopPhrase "Copies only"
        searchPhrase1 = "location" '=> found
        searchPhrase2 = "lizard" '=> not found

    .it "should return True if the search phrase is present"
        cp.SetSearchPhrase searchPhrase1
        .AssertEqual cp.GetResult, True

    .it "should return False if the search phrase is not present"
        cp.SetSearchPhrase searchPhrase2
        .AssertEqual cp.GetResult, False

End With
