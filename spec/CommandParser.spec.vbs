
'test the CommandParser class

With CreateObject("includer")
    Execute(.read("CommandParser"))
    Execute(.read("TestingFramework"))
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
