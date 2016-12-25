
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
        Dim expected, actual

    .it "should copy & get text to & from the clipboard, with spaces"
        expected = "  " & fso.GetTempName & " . "
        cb.SetClipboardText expected
        actual = cb.GetClipboardText
        .AssertEqual actual, expected

    .it "should clear the clipboard"
        expected = ""
        cb.SetClipboardText expected
        actual = cb.GetClipboardText
        .AssertEqual actual, expected

    .it "should copy the word ""off"" to the clipboard"
        expected = "ofF"
        cb.SetClipboardText expected
        actual = cb.GetClipboardText
        .AssertEqual actual, expected

End With

'teardown

Set fso = Nothing
Set sh = Nothing
Set HtmlFile = Nothing
