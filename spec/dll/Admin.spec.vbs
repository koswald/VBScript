
'test Admin.dll without elevated privileges

Option Explicit : Initialize

With New TestingFramework

    .describe "Admin.dll non-elevated privileges"
        Set va = CreateObject("VBScripting.Admin")

    .it "should indicate that privileges are not elevated"
        .AssertEqual va.PrivilegesAreElevated, False

    .it "should verify that an EventLog source is installed"
        .AssertEqual va.SourceExists(source), True

    .it "should raise an error on new-source SourceExists call" 'because of low privileges
        On Error Resume Next
            Dim result : result = va.SourceExists("VBScripting2")
            .AssertEqual Err.Description, "The source was not found, but some or all event logs could not be searched.  Inaccessible logs: Security."
        On Error Goto 0

    .it "should raise an error on new-source CreateEventSource call"
        On Error Resume Next
            result = va.CreateEventSource("VBScripting2")
            .AssertEqual Err.Description, "The source was not found, but some or all event logs could not be searched.  Inaccessible logs: Security."
        On Error Goto 0

    .it "should indicate a known source on CreateEventSource call"
        result = va.CreateEventSource("WSH")
        .AssertEqual result(1), va.Result.SourceAlreadyExists

    .it "should raise an error on deleting non-existent source"
        On Error Resume Next
            result = va.DeleteEventSource("VBScripting2")
            .AssertEqual Err.Description, "The source was not found, but some or all event logs could not be searched.  Inaccessible logs: Security."
        On Error Goto 0

    .it "should raise an error attempting to delete an existing source"
        On Error Resume Next
            result = va.DeleteEventSource("VBScripting")
        .AssertEqual Err.Description, "Requested registry access is not allowed."
        On Error Goto 0

    .it "should read from the event log"
        Dim guid : guid = gg.Generate
        va.Log "Testing VBScipting.Admin..." & _
            vbLf & "unique search string: " & guid
        Dim logs : logs = va.GetLogs(source, guid)
        If UBound(logs) = -1 Then ShowLogNotFoundMessage
        .AssertEqual logs(0), "Testing VBScipting.Admin..." & _
            vbLf & "unique search string: " & guid
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
