
'test TextStreamer.vbs

With CreateObject("includer")
    Execute(.read("TextStreamer"))
    Execute(.read("TestingFramework"))
    Execute(.read("VBSNatives"))
End With
Dim ts : Set ts = New TextStreamer
Dim n : Set n = New VBSNatives

With New TestingFramework

    .describe "TextStreamer class"

    .it "should expose an instance of the StreamConstants object"

        .AssertEqual ts.sc.tbSystemDefault, -2

    .it "should expose an instance of the Scripting.FileSystemObject object"

        .AssertEqual ts.fso.FileExists(WScript.ScriptFullName), True

    .it "should expose an instance of the WScript.Shell object"

        .AssertEqual ts.sh.ExpandEnvironmentStrings("%SystemRoot%"), ts.sh.ExpandEnvironmentStrings("%WinDir%")

    .it "should default to Ascii format"

        .AssertEqual ts.GetStreamFormat, ts.sc.tbAscii

    .it "should default to create a new file"

        .AssertEqual ts.GetCreateMode, ts.sc.bCreateNew

    .it "should default to append"

        .AssertEqual ts.GetStreamMode, ts.sc.iForAppending

    .it "should open a file for appending and for reading"

        Dim sentence : sentence = "free speech is under attack"
        Dim stream : Set stream = ts.Open
        stream.WriteLine sentence
        stream.Close
        ts.SetForReading
        Set stream = ts.Open
        Dim line : line = stream.ReadLine
        stream.Close

        .AssertEqual line, sentence

         ts.Delete

   .it "should open a file for writing"

        ts.SetForWriting
        Set stream = ts.Open
        stream.WriteLine sentence
        stream.Close
        ts.SetForReading
        Set stream = ts.Open
        line = stream.ReadLine
        stream.Close

        .AssertEqual line, sentence

        ts.Delete

End With
