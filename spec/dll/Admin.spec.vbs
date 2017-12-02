
'test Admin.dll without elevated privileges

Option Explicit : Initialize

With New TestingFramework

    .describe "Admin.dll non-elevated privileges"
        Set va = CreateObject("VBScripting.Admin")

    .it "should verify that an EventLog source is installed"
        .AssertEqual va.SourceExists(source), True

    .it "should show message on new-source SourceExists call"
        sh.Run "fixture\AdminNonExistentSourceExists.vbs"
        .AssertEqual .MessageAppeared( _
            "Failed to find event source", .5, "{Enter}"), True

    .it "should show message on new-source CreateEventSource call"
        sh.Run "fixture\AdminCreateNonExistentSource.vbs"
        .AssertEqual .MessageAppeared( _
            "Failed to find event source", .5, "{Enter}"), True

    .it "should show message on known-source CreateEventSource call"
        sh.Run "fixture\AdminCreateExistingSource.vbs"
        .AssertEqual .MessageAppeared( _
            "Source exists", .5, "{Enter}"), True

    .it "should show message on deleting non-existent source"
        sh.Run "fixture\AdminDeleteNonExistentSource.vbs"
        .AssertEqual .MessageAppeared( _
            "Failed to find event source", .5, "{Enter}"), True

    .it "should show message on deleting an existing source"
        sh.Run "fixture\AdminDeleteExistingSource.vbs"
        .AssertEqual .MessageAppeared( _
            "Couldn't delete source", .5, "{Enter}"), True

    .it "should read from the event log"
        Dim guid : guid = gg.Generate
        va.Log "Testing VBScipting.Admin..." & _
            vbLf & "unique search string: " & guid
        Dim logs : logs = va.GetLogs(source, guid)
        If UBound(logs) = -1 Then ShowLogNotFoundMessage
        .AssertEqual logs(0), "Testing VBScipting.Admin..." & _
            vbLf & "unique search string: " & guid

    .it "should indicate that privileges are not elevated"
        .AssertEqual va.PrivilegesAreElevated, False
End With

Quit

Sub Quit
    ReleaseObjectMemory
    WScript.Quit
End Sub

Sub ReleaseObjectMemory
    Set va = Nothing
    Set sh = Nothing
End Sub

Sub ShowLogNotFoundMessage
    Dim msg : msg = vbLf & _
        "Couldn't find the log. Do you have the source " & vbLf & _
        "configured correctly in " & configFile & "? " & vbLf & _
        "Current value: " & source & vbLf
    WScript.StdOut.WriteLine msg
End Sub

Dim va, sh, log, gg
Dim source
Const configFile = "../spec/dll/Admin.spec.config"

Sub Initialize
    With CreateObject("includer")
        Execute .read("PrivilegeChecker")
        Execute .read("VBSEventLogger")
        Execute .read("GuidGenerator")
        ExecuteGlobal .read("TestingFramework")
        Execute .read(configFile)
    End With
    Dim pc : Set pc = New PrivilegeChecker
    Set gg = New GuidGenerator
    Set log = New VBSEventLogger
    'log 4, "logger test"
    If pc Then Err.Raise 1,, WScript.ScriptName & " requires that privileges not be elevated."
    Set sh = CreateObject("WScript.Shell")
End Sub
