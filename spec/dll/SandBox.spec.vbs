
'test SandBox.dll

Option Explicit : Initialize

With New TestingFramework
    
    .describe "SandBox.dll"

    .it "should get an AdHoc int"
        .AssertEqual sb.AdHoc, 1

End With

Cleanup

Sub Cleanup
    Set sb = Nothing
    Set fso = Nothing
    Set sh = Nothing
End Sub

Dim sb, fso, sh, log

Sub Initialize
    Set sb = CreateObject("VBScripting.SandBox")
    With CreateObject("includer")
        Execute .read("VBSEventLogger")
        ExecuteGlobal .read("TestingFramework")
    End With
    Set log = New VBSEventLogger
'    log 1, WScript.ScriptName & ": error entry test"
'    log 2, WScript.ScriptName & ": warning entry test"
'    log 4, WScript.ScriptName & ": information entry test"
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

