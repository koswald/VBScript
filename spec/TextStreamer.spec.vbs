
'test TextStreamer.vbs

With CreateObject("includer")
    Execute .read("TextStreamer")
    Execute .read("TestingFramework")
    Execute .read("StreamConstants")
End With
Dim ts : Set ts = New TextStreamer
Dim sc : Set sc = New StreamConstants
Dim sh : Set sh = CreateObject("WScript.Shell")
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

With New TestingFramework

    .describe "TextStreamer class"

    .it "should expose an instance of the StreamConstants object"

        .AssertEqual sc.tbSystemDefault, -2

    .it "should expose an instance of the Scripting.FileSystemObject object"

        .AssertEqual fso.FileExists(WScript.ScriptFullName), True

    .it "should expose an instance of the WScript.Shell object"

        .AssertEqual sh.ExpandEnvironmentStrings("%SystemRoot%"), sh.ExpandEnvironmentStrings("%WinDir%")

    .it "should default to Ascii format"

        .AssertEqual ts.GetStreamFormat, sc.tbAscii

    .it "should default to create a new file"

        .AssertEqual ts.GetCreateMode, sc.bCreateNew

    .it "should default to append"

        .AssertEqual ts.GetStreamMode, sc.iForAppending

    .it "should open a file for appending and for reading"

        Dim sentence : sentence = "free speech is under attack"
        Dim stream : Set stream = ts.Open
        stream.WriteLine sentence
        stream.Close
        ts.SetForReading
        Set stream = ts.Open

        .AssertEqual stream.ReadLine, sentence

        stream.Close
        ts.Delete

   .it "should open a file for writing"

        ts.SetForWriting
        Set stream = ts.Open
        stream.WriteLine sentence
        stream.Close
        ts.SetForReading
        Set stream = ts.Open

        .AssertEqual stream.ReadLine, sentence

        stream.Close
        ts.Delete

    'garbage collection
        Set sh = Nothing
        Set fso = Nothing

End With
