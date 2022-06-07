'Test Admin.dll with elevated privileges

Option Explicit : Initialize

With New TestingFramework

    .Describe "VBScripting.Admin object - elevated"

    .It "should indicate privileges are elevated"
        .AssertEqual va.PrivilegesAreElevated, True

    .It "should verify that an EventLog source is installed"
        .AssertEqual va.SourceExists(va.EventSource), True

    .It "should indicate if a source already exists on CreateEventSource call"
        Set result = va.CreateEventSource( "WSH" )
        .AssertEqual result.Result, va.Result.SourceAlreadyExists

    .It "should indicate if a source doesn't exist on DeleteEventSource call"
        Set result = va.DeleteEventSource( va.EventSource & "2" )
        .AssertEqual result.Result, va.Result.SourceDoesNotExist

    .It "should indicate success after creating a source"
        Set result = va.CreateEventSource( va.EventSource & "2" )
        .AssertEqual result.Result, va.Result.SourceCreated

    .It "should indicate success after deleting a source"
        Set result = va.DeleteEventSource( va.EventSource & "2" )
        .AssertEqual result.Result, va.Result.SourceDeleted
End With

Quit

Sub Cleanup
    If va.SourceExists( va.EventSource & "2" ) Then
        va.DeleteEventSource va.EventSource & "2"
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
    Set result = Nothing
End Sub

Const elevated = True
Dim va, sh, log
Dim result

Sub Initialize
    With CreateObject( "VBScripting.Includer" )
        Execute .Read( "PrivilegeChecker" )
        Execute .Read( "VBSEventLogger" )
        Execute .Read( "VBSApp" )
        Execute .Read( "TestingFramework" )
    End With
    Dim pc : Set pc = New PrivilegeChecker
    Set log = New VBSEventLogger
    Dim app : Set app = New VBSApp
    'log 4, "logger test"
    If Not pc Then
        app.SetUserInteractive False
        app.RestartWith app.CScriptHost, "/k", elevated
    End If
    Set sh = CreateObject( "WScript.Shell" )
    Set va = CreateObject( "VBScripting.Admin" )
    Cleanup
End Sub
