
'clipboard procedures

Class VBSClipboard

	Private sh, hidden, synchronous

	Sub Class_Initialize
		Set sh = CreateObject("WScript.Shell")
        hidden = 0
        synchronous = True
	End Sub

	'Method SetClipText
	'Parameter: a string
	'Remark: Copies the specified string to the clipboard. Uses clip.exe, which shipped with Windows since Vista / Windows Server 2003.

	Sub SetClipText(newText)
		sh.Run "cmd.exe /c echo " & newText & " | clip", hidden, synchronous
	End Sub

    'Undocumented Function TrimTheTail
    'Remark: Trims the spurious characters added to the clipboard text by the htmlfile object. Used internally and by the unit test

    Function TrimTheTail(text_)
        Dim text : text = text_
        Dim tail : tail = Array(32, 13, 10) 'space, vbCr, vbLf: characters added by the htmlfile obj
        Dim i
        For i = UBound(tail) To 0 Step -1
            If Right(text, 1) = Chr(tail(i)) Then
                text = Left(text, Len(text) - 1)
            End If
        Next
        TrimTheTail = text
    End Function

	Sub Class_Terminate
		Set sh = Nothing
	End Sub

End Class