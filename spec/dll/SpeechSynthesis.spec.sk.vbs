'test SpeechSynthesis.dll

Option Explicit : Initialize

With New TestingFramework

    .describe "SpeechSynthesis.dll - warning: test uses SendKeys!"
        Set ss = CreateObject("VBScripting.SpeechSynthesis")
        
    .ShowSendKeysWarning

    .it "should show a message on setting an invalid voice"
        sh.Run "fixture\SpeechSynthesis.invalid-voice.vbs"
        .AssertEqual .MessageAppeared( _
            "SpeechSynthesis class", 2, "{Enter}"), True

    .CloseSendKeysWarning   
End With

Cleanup

Dim ss, sh, i
Const asynchronous = False

Sub Initialize
    With CreateObject("includer")
        ExecuteGlobal .read("TestingFramework")
    End With
    Set sh = CreateObject("WScript.Shell")
End Sub

Sub Cleanup
    ss.Dispose
    Set ss = Nothing
    Set sh = Nothing
End Sub
