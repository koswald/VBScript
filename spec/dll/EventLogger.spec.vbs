'Test EventLogger.dll

Option Explicit
Dim el 'VBScripting.EventLogger object under test
Dim ad 'VBScripting.Admin object
Dim incl 'VBScripting.Includer object
Dim gg 'GuidGenerator object
Dim guid 'a unique string to be searched
Dim logs 'array of strings returned by GetLogs

Set ad = CreateObject( "VBScripting.Admin" )
Set incl = CreateObject( "VBScripting.Includer" )
Execute incl.Read( "GuidGenerator" )
Set gg = New GuidGenerator
guid = gg.Generate

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "VBScripting.EventLogger object"
        Set el = CreateObject( "VBScripting.EventLogger" )

    .It "should write to the event log"
        .WriteTempMessage "Writing to the event logs..."
        el.log "Testing EventLogger.dll... " & guid
        .EraseTempMessage
        .WriteTempMessage "Reading from the event logs..."
        logs = ad.GetLogs("VBScripting", guid)
        .EraseTempMessage
        .AssertEqual "Testing EventLogger.dll... " & guid, logs(0)

End With

