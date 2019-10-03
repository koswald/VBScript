
'test ValidFileName.vbs

Initialize
Main
Set includer = Nothing

Sub Main
    With New TestingFramework

        .describe "ValidFileName.vbs"

        .it "should raise an error when Execute is used" 'to instantiate ValidFileName.vbs functions when scope is not global
            Execute includer.read("ValidFileName")
            On Error Resume Next
                Dim x : x = GetValidFileName("xx")
                .AssertEqual Err.Description, "Use ExecuteGlobal, not Execute, with Function-based scripts like ValidFileName.vbs"
            On Error Goto 0

        .it "should return a string suitable for a filename"
            ExecuteGlobal includer.read("ValidFileName")
            .AssertEqual GetValidFileName("\/:*?""<>|%20#"), "-----------"

    End With
End Sub

Dim includer
Sub Initialize
    Set includer = CreateObject("VBScripting.Includer")
    ExecuteGlobal includer.read("TestingFramework")
End Sub
