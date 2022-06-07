'Test Admin.dll without elevated privileges

Option Explicit : Initialize

With New TestingFramework

    .Describe "VBScripting.Admin object - non-elevated"
        Set va = CreateObject( "VBScripting.Admin" )

    .It "should indicate that privileges are not elevated"
        .AssertEqual va.PrivilegesAreElevated, False

    .It "should verify that an EventLog source is installed"
        .AssertEqual va.SourceExists(va.EventSource), True

    .It "should raise an error on new-source SourceExists call" 'because of low privileges
        On Error Resume Next
            Dim result : result = va.SourceExists( va.EventSource & "2" )
            .AssertEqual Left(Err.Description, 75), "The source was not found, but some or all event logs could not be searched."
        On Error Goto 0

    .It "should raise an error on new-source CreateEventSource call"
        On Error Resume Next
            Set result = va.CreateEventSource( va.EventSource & "2" )
            .AssertEqual Left(Err.Description, 75), "The source was not found, but some or all event logs could not be searched."
        On Error Goto 0

    .It "should indicate a known source on CreateEventSource call"
        Set result = va.CreateEventSource( "WSH" )
        .AssertEqual result.Result, va.Result.SourceAlreadyExists

    .It "should raise an error on deleting non-existent source"
        On Error Resume Next
            result = va.DeleteEventSource( va.EventSource & "2" )
            .AssertEqual Left(Err.Description, 75), "The source was not found, but some or all event logs could not be searched."
        On Error Goto 0

    .It "should fail to delete an existing source without elevated privileges"
        Set result = va.DeleteEventSource(va.EventSource)
        .AssertEqual result.Result, va.Result.SourceDeletionException

    .It "should return a SourceExists boolean in the result"
        .AssertEqual TypeName(result.SourceExists), "Boolean"
        .ShowPendingResult 'help to clarify which spec is causing the delay (reading from the log below)

    .It "should read from the event log"
        Dim guid : guid = gg.Generate
        .WriteTempMessage "Writing to the event logs..."
        va.Log "Testing VBScipting.Admin..." & vbLf & _
            "unique search string: " & guid
        .EraseTempMessage
        .WriteTempMessage "Reading from the event logs..."
        Dim logs : logs = va.GetLogs(va.EventSource, guid)
        .EraseTempMessage
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
        "Couldn't find the log." & vbLf
    WScript.StdOut.WriteLine msg
End Sub

Dim va, sh, log, gg

Sub Initialize
    With CreateObject( "VBScripting.Includer" )
        Execute .Read( "PrivilegeChecker" )
        Execute .Read( "VBSEventLogger" )
        Execute .Read( "GuidGenerator" )
        ExecuteGlobal .Read( "TestingFramework" )
    End With
    Dim pc : Set pc = New PrivilegeChecker
    Set gg = New GuidGenerator
    Set log = New VBSEventLogger
    'log 4, "logger test"
    If pc Then Err.Raise 17,, WScript.ScriptName & " requires that privileges not be elevated."
    Set sh = CreateObject( "WScript.Shell" )
End Sub
