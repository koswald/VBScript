'Test VBSClipboard.vbs

Option Explicit
Dim cb 'VBSClipboard object, under test
Dim fso 'Scripting.FileSystemObject object
Dim incl 'VBScripting.Includer object
Dim expected, actual 'assertion arguments

Set fso = CreateObject( "Scripting.FileSystemObject" )
Set incl = CreateObject( "VBScripting.Includer" )
Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .describe "VBSClipboard class"
        Execute incl.Read( "VBSClipboard" )
        Set cb = New VBSClipboard

    'setup

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

Set fso = Nothing
