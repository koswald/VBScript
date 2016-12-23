
'test VBSClipboard.vbs

With CreateObject("includer")
    Execute(.read("VBSClipboard"))
    Execute(.read("TestingFramework"))
    Execute(.read("VBSLogger"))
End With

With New TestingFramework

    .describe "VBSClipboard class"
        Dim cb : Set cb = New VBSClipboard

    'setup
        Const hidden = 0, synchronous = True 'WScript.Shell Run method constants
        Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim sh : Set sh = CreateObject("WScript.Shell")
        Dim HtmlFile : Set HtmlFile = CreateObject("htmlfile")
        Dim log : Set log = New VBSLogger
        Dim randomText
        Dim DiscrepancyFound : DiscrepancyFound = False
        Dim actual, expected

    .it "should copy text to the clipboard"
        randomText = fso.GetTempName
        cb.SetClipText randomText
        .AssertEqual cb.TrimHtmlFileData(HtmlFile.parentWindow.ClipboardData.GetData("text")), randomText

    .it "should get text from the clipboard"
        randomText = fso.GetTempName
        'copy the test text to the clipboard
        sh.Run "cmd.exe /c echo " & randomText & " | clip", hidden, synchronous

        expected = randomText
        actual = cb.GetClipText
        .AssertEqual actual, expected

        LogAnyDiscrepency

    .it "should clear the clipboard"
        cb.SetClipText ""
        .AssertEqual cb.TrimHtmlFileData(HtmlFile.parentWindow.ClipboardData.GetData("text")), ""

    .it "should copy the word ""off"" to the clipboard"
        cb.SetClipText "ofF"
        .AssertEqual LCase(cb.TrimHtmlFileData(HtmlFile.parentWindow.ClipboardData.GetData("text"))), "off"
End With

'teardown

Set fso = Nothing
Set sh = Nothing
Set HtmlFile = Nothing

If DiscrepancyFound Then log.View 'open the log file for viewing (default=Notepad)

'Log the Ascii code for each "actual" character

Sub LogAnyDiscrepency
    If actual = expected Then Exit Sub
    DiscrepancyFound = True
    Dim i, a : a = actual
    Dim d : d = "Characters: " 'discrepancy
    For i = 1 To Len(a)
        d = d & Asc(Left(a, 1)) & " "
        a = Right(a, Len(a) - 1)
    Next
    log d
End Sub
