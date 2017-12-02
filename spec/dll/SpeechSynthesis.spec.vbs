'test SpeechSynthesis.dll

Option Explicit : Initialize

With New TestingFramework

    .describe "Speech synthesis library for VBScript"
        Set ss = CreateObject("VBScripting.SpeechSynthesis")
        
    .it "should get the volume"
        On Error Resume Next
            Dim vol : vol = ss.Volume
            .AssertEqual Err.Description, ""
        On Error Goto 0
        
    .it "should set the volume"
        On Error Resume Next
            ss.Volume = 0
            .AssertEqual Err.Description, ""
        On Error Goto 0
        
    .it "should have a Speak method"
        On Error Resume Next
             ss.Speak "test"
             .AssertEqual Err.Description, ""
        On Error Goto 0
        
    .it "should indicate when ready for SpeakAsync call"
         .AssertEqual ss.SynthesizerState, ss.State.Ready
    
    .it "should have a SpeakAsync method"
        On Error Resume Next
            ss.SpeakAsync "1234"
            .AssertEqual Err.Description, ""
        On Error Goto 0
        
    .it "should pause asynchronous speech"
        On Error Resume Next
            WScript.Sleep 50 'wait for state change
            ss.Pause
            .AssertEqual ss.SynthesizerState, ss.State.Paused
        On Error Goto 0
         
    .it "should resume paused speech"
        WScript.Sleep 50 'wait for state change
        ss.Resume
        .AssertEqual ss.SynthesizerState, ss.State.Speaking
 
    .it "should return a list of installed, enabled voices:"
        Dim voices : voices = ss.Voices
        .AssertEqual TypeName(voices), "Variant()"
        .ShowPendingResult
        For i = 0 To UBound(voices)
            WScript.StdOut.WriteLine "              " & voices(i)
        Next

    .it "should get the current voice"
        On Error Resume Next
            Dim voice : voice = ss.Voice
            .AssertEqual Err.Description, ""
        On Error Goto 0
        
    .it "should set the current voice"
        On Error Resume Next
            ss.Voice = voices(0)
            .AssertEqual Err.Description, ""
        On Error Goto 0
        
    .it "should show a message on setting an invalid voice"
        sh.Run "fixture\SpeechSynthesis.invalid-voice.vbs"
        .AssertEqual .MessageAppeared( _
            "SpeechSynthesis class", 2, "{Enter}"), True
   
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
