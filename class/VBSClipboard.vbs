
'Clipboard procedures

Class VBSClipboard

    Private hidden, synchronous
    Private sh, HtmlFile, formatter
    Private TextToCut, ClipExe

    Sub Class_Initialize
        hidden = 0 : synchronous = True 'WScript.Shell Run method constants
        Set sh = CreateObject("WScript.Shell")
        Set HtmlFile = CreateObject("htmlfile")
        With CreateObject("includer")
            Execute(.read("StringFormatter"))
        End With
        Set formatter = New StringFormatter
        ClipExe = "%SystemRoot%\System32\clip.exe"
        TextToCut = Chr(32) & Chr(13) & Chr(10) 'space, vbCr, vbLf: extra characters added by the htmlfile obj
    End Sub

    'Method SetClipboardText
    'Parameter: a string
    'Remark: Copies the specified string to the clipboard. Uses clip.exe, which shipped with Windows&reg; Vista / Server 2003 through Windows 10.

    Sub SetClipboardText(ByVal newText)
        If newText = "" Then
            newText = "off" ' "echo off | clip" clears the clipboard
        ElseIf Trim(LCase(newText)) = "off" Then
            SetClipboardTextAlt newText
            Exit Sub
        End If
        sh.Run formatter(Array("cmd /c echo %s | %s", newText, ClipExe)), hidden, synchronous
    End Sub

    'Private method SetClipboardTextAlt
    'Slower alternate method intended for when Trim(LCase(TheTextToCopyToTheClipboard)) = "off"

    Private Sub SetClipboardTextAlt(newText)
        'create a temp file with contents equal to newText,
        'then send the contents to clip.exe
        With CreateObject("includer") : On Error Resume Next
            Execute(.read("TextStreamer"))
        End With : On Error Goto 0
        Dim ts : Set ts = New TextStreamer
        Dim stream : Set stream = ts.Open
        stream.Write newText
        stream.Close
        Set stream = Nothing
        sh.Run formatter(Array("cmd /c type ""%s"" | %s", ts.GetFile, ClipExe)), hidden, synchronous
        ts.Delete
    End Sub

    'Property GetClipboardText
    'Returns a string
    'Remark: Returns text from the clipboard

    Property Get GetClipboardText
        Const MaxTries = 5
        Dim tries : tries = 0
        On Error Resume Next
            While Err.Number <> 0 Or GetClipboardText = TextToCut Or TypeName(GetClipboardText) = "Null" Or tries = 0
                Err.Clear
                GetClipboardText = TrimHtmlFileData(HtmlFile.parentWindow.ClipboardData.GetData("text"))
                tries = tries + 1
                If tries > MaxTries Then Err.Raise 1, WScript.ScriptName, "VBSClipboard.GetClipboardText failed to get the clipboard text after " & MaxTries & " tries."
            Wend
        On Error Goto 0
    End Property

    'Private Function TrimHtmlFileData
    'Trims the spurious characters added to the clipboard text by the htmlfile object. Used internally and by the unit test.

    Private Function TrimHtmlFileData(ByVal text)
        Dim k : k = Len(TextToCut)
        If Right(text, k) = TextToCut Then text = Left(text, Len(text) - k)
        TrimHtmlFileData = text
    End Function

    Sub Class_Terminate
        Set sh = Nothing
        Set HtmlFile = Nothing
    End Sub

End Class

