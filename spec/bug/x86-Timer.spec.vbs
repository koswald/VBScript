Option Explicit : Setup
With New TestingFramework

    .describe "VBScripting.Timer object 32-bit bug test"
        Dim csTimer : Set csTimer = CreateObject( "VBScripting.Timer" )

    .it "should reproduce error/bug"
        On Error Resume Next
            x = csTimer.Interval/1
            .AssertEqual Err.Description, _
               "Variable uses an Automation type not supported in VBScript"
        On Error Goto 0

    .it "should demonstrate workaround syntax: convert to CLng"
        On Error Resume Next
            x = CLng(csTimer.Interval)/1
            .AssertEqual Err.Description, ""
        On Error Goto 0

End With
Teardown

Dim x
Dim sh, fso, includer, wow

Sub Setup
    Dim cscriptX86
    cscriptX86 = "%SystemRoot%\SysWoW64\cscript.exe"
    Set sh = CreateObject( "WScript.Shell" )
    Set fso = CreateObject( "Scripting.FileSystemObject" )
    Set includer = CreateObject( "VBScripting.Includer" )
    ExecuteGlobal includer.Read( "TestingFramework" )
    Execute includer.Read( "WowChecker" )
    Set wow = New WowChecker
    Execute includer.Read( "VBSApp" )
    With New VBSApp
        If "cscript.exe" = .GetExe _
        And wow.IsWow Then
            WScript.StdOut.WriteLine "using the 32-bit cscript.exe..."
            Exit Sub
        End If
        .RestartUsing cscriptX86, .DoNotExit, .DoNotElevate
    End With
End Sub
Sub Teardown
    Set sh = Nothing
    Set fso = Nothing
    Set wow = Nothing
    Set includer = Nothing
    Set csTimer = Nothing
End Sub
