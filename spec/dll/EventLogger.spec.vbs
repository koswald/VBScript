
'test EventLogger.dll

Option Explicit : Initialize

With New TestingFramework
    
    .describe "EventLogger.dll"

    .it "should write to the event log"
        Dim guid : guid = gg.Generate
        el.log "Testing EventLogger.dll... " & guid
        Dim logs : logs = ad.GetLogs("VBScripting", guid)
        .AssertEqual "Testing EventLogger.dll... " & guid, logs(0)

End With

Dim el, ad, gg, sh

Sub Initialize
    Set el = CreateObject("VBScripting.EventLogger")
    Set ad = CreateObject("VBScripting.Admin")
    With CreateObject("includer")
        Execute .read("GuidGenerator")
        ExecuteGlobal .read("TestingFramework")
    End With
    Set gg = New GuidGenerator
    Set sh = CreateObject("WScript.Shell")
End Sub

