Option Explicit : Setup
With New TestingFramework

    .describe "VBScripting.Timer class"
        Set vt = CreateObject("VBScripting.Timer")
    .it "should return a Single type"
        vt.IntervalInHours = 3
        .AssertEqual TypeName(vt.IntervalInHours), "Single"
    .it "should return the expected Single"
        vt.IntervalInHours = 1.54
        .AssertEqual vt.IntervalInHours, 1.54

End With
Teardown

Dim vt, sh, fso, includer
Sub Setup
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set includer = CreateObject("VBScripting.Includer")
    ExecuteGlobal includer.Read("TestingFramework")
End Sub
Sub Teardown
    Set vt = Nothing
    Set sh = Nothing
    Set fso = Nothing
    Set includer = Nothing
End Sub
