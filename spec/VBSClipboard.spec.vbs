
'test VBSClipboard.vbs

With CreateObject("includer")
    Execute(.read("VBSClipboard"))
    Execute(.read("TestingFramework"))
End With

With New TestingFramework

    .describe "VBSClipboard class"
        Dim cb : Set cb = New VBSClipboard

    'setup
        Const hidden = 0, synchronous = True 'WScript.Shell Run method constants
        Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim sh : Set sh = CreateObject("WScript.Shell")
        Dim HtmlFile : Set HtmlFile = CreateObject("htmlfile")
        Dim randomText
        Dim actual, expected

    .it "should copy text to the clipboard"
        randomText = fso.GetTempName
        cb.SetClipText randomText
        .AssertEqual cb.TrimHtmlFileData(HtmlFile.parentWindow.ClipboardData.GetData("text")), randomText

    .it "should get text from the clipboard"
        randomText = fso.GetTempName

        expected = randomText

        sh.Run "cmd.exe /c echo " & randomText & " | clip", hidden, synchronous 'set clipboard text
        '.AssertEqual cb.GetClipText, randomText

        actual = cb.GetClipText
        ShowDiscrepency
        .AssertEqual actual, expected
End With

'teardown

Set fso = Nothing
Set sh = Nothing
Set HtmlFile = Nothing

Sub ShowDiscrepency
    If actual = expected Then Exit Sub
    Dim i, a : a = actual
    WScript.StdOut.Write "Characters: "
    For i = 1 To Len(a)
        WScript.StdOut.Write Asc(Left(a, 1)) & " "
        a = Right(a, Len(a) - 1)
    Next
    WScript.StdOut.WriteLine ""
End Sub
