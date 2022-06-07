'Test SpeechSynthesis.dll

Option Explicit : Initialize

With New TestingFramework

    .Describe "VBScripting.SpeechSynthesis object"
        Set ss = CreateObject( "VBScripting.SpeechSynthesis" )

    .It "should get the volume"
        On Error Resume Next
            Dim vol : vol = ss.Volume
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should set the volume"
        On Error Resume Next
            ss.Volume = 0
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should have a Speak method"
        On Error Resume Next
             ss.Speak "test"
             .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should indicate when ready for SpeakAsync call"
         .AssertEqual ss.SynthesizerState, ss.State.Ready

    .It "should have a SpeakAsync method"
        On Error Resume Next
            ss.SpeakAsync "1234"
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should pause asynchronous speech"
        On Error Resume Next
            WScript.Sleep 50 'wait for state change
            ss.Pause
            .AssertEqual ss.SynthesizerState, ss.State.Paused
        On Error Goto 0

    .It "should resume paused speech"
        WScript.Sleep 50 'wait for state change
        ss.Resume
        .AssertEqual ss.SynthesizerState, ss.State.Speaking

    .It "should return a list of installed, enabled voices:"
        Dim voices : voices = ss.Voices
        .AssertEqual TypeName(voices), "Variant()"
        .ShowPendingResult
        For i = 0 To UBound(voices)
            WScript.StdOut.WriteLine "              " & voices(i)
        Next

    .It "should get the current voice"
        On Error Resume Next
            Dim voice : voice = ss.Voice
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should set the current voice"
        On Error Resume Next
            ss.Voice = voices(0)
            .AssertEqual Err.Description, ""
        On Error Goto 0
        
End With

Cleanup

Dim ss, sh, i
Const asynchronous = False

Sub Initialize
    With CreateObject( "VBScripting.Includer" )
        ExecuteGlobal .Read( "TestingFramework" )
    End With
    Set sh = CreateObject( "WScript.Shell" )
End Sub

Sub Cleanup
    ss.Dispose
    Set ss = Nothing
    Set sh = Nothing
End Sub
