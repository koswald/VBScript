
'script for SpeechSynthesis.hta

'convert text to sound
Sub Speak
    ss.SpeakAsync words.value
    words.select 'select all; prepare to overwrite the previous words
End Sub

'event handler
Sub KeyUp
    If EnterKey = window.event.keyCode Then
        Speak
    End If
End Sub

'event handler
Sub ChangeVoice
    voiceIndex = voiceIndex + 1
    If voiceIndex > UBound(voices) Then voiceIndex = 0
    ss.Voice = voices(voiceIndex)
    voiceButton.Title = "Current voice: " & voices(voiceIndex)
    words.select
End Sub

Const width = 30, height = 30 'window size in percent of screen
Const xPos = 50, yPos = 50 'window position in percent of screen
Const EnterKey = 13
Dim words, voiceButton 'html elements
Dim ss 'speech synthesis object
Dim voices, voiceIndex, nVoices

'initialize html elements and the ss object
Sub Window_OnLoad
    Dim application : Set application = document.getElementsByTagName("application")(0)
    document.title = application.ApplicationName

    'set window size and position
    Dim pxWidth, pxHeight 'window size in pixels
    With document.parentWindow.screen
        pxWidth = .availWidth * width * .01
        pxHeight = .availHeight * height * .01
        self.ResizeTo pxWidth, pxHeight
        self.MoveTo _
            (.availWidth - pxWidth) * xPos * .01005, _
            (.availHeight - pxHeight) * yPos * .0102
    End With

    'create a container for the button
    Dim ctnr1 : Set ctnr1 = document.createElement("div")
    With ctnr1.style
        .width = "100%" 'keep the button above the text area
        .height = "20%"
    End With
    document.body.insertBefore ctnr1

    'create the Speak button
    Dim button : Set button = document.createElement("input")
    With button
        .type = "button"
        .value = "Speak"
        Set .onClick = GetRef("Speak")
        With .style
            .marginBottom = ".5em"
            .height = "80%"
            .width = "30%"
        End With
    End With
    ctnr1.insertBefore button

    'create the change voice button
    Set voiceButton = document.createElement("input")
    With voiceButton
        .type = "button"
        .value = "Change voice"
        Set .onClick = GetRef("ChangeVoice")
        With .style
            .marginBottom = ".5em"
            .height = "80%"
            .width = "40%"
            .marginLeft = "10%"
        End With
    End With
    ctnr1.insertBefore voiceButton

    'create a container for the text area
    Dim ctnr2 : Set ctnr2 = document.createElement("div")
    With ctnr2.style
        .marginBottom = 0
        .height = "80%"
    End With
    document.body.insertBefore ctnr2

    'create the text area
    Set words = document.createElement("textarea")
    Set words.onKeyUp = GetRef("KeyUp")
    With words.style
        .width = "100%"
        .height = "100%"
        .fontFamily = "Comic Sans MS, arial, sans-serif"
        .fontWeight = "bold"
        .fontSize = "larger"
        .overflow = "auto"
    End With
    ctnr2.insertBefore words

    'get the speech synthesizer
    On Error Resume Next
        Set ss = CreateObject("VBScripting.SpeechSynthesis")
        If Err Then
            MsgBox "Failed to find or initiailize the SpeechSynthesis library.", vbCritical, "Initialization failure"
            Self.close
        End If
    On Error Goto 0
    voices = ss.Voices
    voiceIndex = 0
    nVoices = UBound(voices) + 1 'number of installed, enabled voices
    If nVoices > 0 Then
        ss.Voice = voices(voiceIndex)
        voiceButton.Title = "Current voice: " & voices(voiceIndex)
    End If
    If nVoices < 2 Then voiceButton.disabled = True

    'put focus on the text area; prepare to type some words
    words.focus
End Sub

Sub Window_OnUnload
    Set ss = Nothing
    Set words = Nothing
End Sub
