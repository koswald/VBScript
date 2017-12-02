
'test Admin.dll with elevated privileges

Option Explicit : Initialize

With New TestingFramework

    .describe "Admin.dll elevated privileges"

    .it "should verify that an EventLog source is installed"
        .AssertEqual va.SourceExists(source), True

    .it "should show error message on known-source CreateEventSource call"
        sh.Run "fixture\AdminCreateExistingSource.vbs"
        .AssertEqual .MessageAppeared("Source exists", .5, "{Enter}"), True

    .it "should show a success message after creating a source"
        sh.Run "fixture\AdminCreateNonExistentSource.vbs"
        .AssertEqual .MessageAppeared("Source created", .5, "{Enter}"), True

    .it "should indicate privileges are elevated"
        .AssertEqual va.PrivilegesAreElevated, True
End With

Quit

Sub Cleanup
    If va.SourceExists("VBScripting2") Then
        sh.Run "fixture\AdminDeleteTestSource.vbs"
        Dim tf : Set tf = New TestingFramework
        Dim x : x = tf.MessageAppeared("Source deleted", 1, "{Enter}")
    End If
End Sub

Sub Quit
    Cleanup
    ReleaseObjectMemory
    WScript.Quit
End Sub

Sub ReleaseObjectMemory
    Set va = Nothing
    Set sh = Nothing
End Sub

Dim va, sh, log
Dim source 'initialized from Admin.spec.config
Const elevated = True

Sub Initialize
    With CreateObject("includer")
        Execute .read("PrivilegeChecker")
        Execute .read("VBSEventLogger")
        Execute .read("VBSApp")
        ExecuteGlobal .read("TestingFramework")
        Execute .read("../spec/dll/Admin.spec.config")
    End With
    Dim pc : Set pc = New PrivilegeChecker
    Set log = New VBSEventLogger
    Dim app : Set app = New VBSApp
    'log 4, "logger test"
    If Not pc Then
        app.SetUserInteractive False
        app.RestartWith app.CScriptHost, "/k", elevated
    End If
    Set sh = CreateObject("WScript.Shell")
    Set va = CreateObject("VBScripting.Admin")
    Cleanup
End Sub
