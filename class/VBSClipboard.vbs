
'Clipboard procedures

Class VBSClipboard

    Private sh, hidden, synchronous
    Private HtmlFile, TextToCut

    Sub Class_Initialize
        hidden = 0 : synchronous = True 'WScript.Shell Run method constants for args 2 & 3
        Set sh = CreateObject("WScript.Shell")
        Set HtmlFile = CreateObject("htmlfile")
        TextToCut = Chr(32) & Chr(13) & Chr(10) 'space, vbCr, vbLf: extra characters added by the htmlfile obj
    End Sub

    'Method SetClipText
    'Parameter: a string
    'Remark: Copies the specified string to the clipboard. Uses clip.exe, which shipped with Windows&reg; Vista / Server 2003 through Windows 10.

    Sub SetClipText(newText)
        sh.Run "cmd.exe /c echo " & newText & " | clip", hidden, synchronous
    End Sub

    'Property GetClipText
    'Returns a string
    'Remark: Returns text from the clipboard
    'ToDo: Loop count limit.
    'ToDo: Test edge case where cliboard text is null; double or triple check null value

    Property Get GetClipText
        GetClipText = ""
        On Error Resume Next
            While Err.Number <> 0 Or GetClipText = "" Or GetClipText = TextToCut Or TypeName(GetClipText) = "Null"
                Err.Clear
                GetClipText = TrimHtmlFileData(HtmlFile.parentWindow.ClipboardData.GetData("text"))
            Wend
        On Error Goto 0
    End Property

    'Undocumented Function TrimHtmlFileData
    'Trims the spurious characters added to the clipboard text by the htmlfile object. Used internally and by the unit test.

    Function TrimHtmlFileData(text_)
        Dim text : text = text_
        Dim k : k = Len(TextToCut)
        If Right(text, k) = TextToCut Then text = Left(text, Len(text) - k)
        TrimHtmlFileData = text
    End Function

    Sub Class_Terminate
        Set sh = Nothing
        Set HtmlFile = Nothing
    End Sub

End Class

