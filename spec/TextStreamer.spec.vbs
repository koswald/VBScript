'test TextStreamer.vbs

Dim incl : Set incl = CreateObject( "VBScripting.Includer" )
Dim sh : Set sh = CreateObject( "WScript.Shell" )
Dim fso : Set fso = CreateObject( "Scripting.FileSystemObject" )
Const Ascii = 0
Const Unicode = -1
Const SystemDefault = -2
Const ForAppending = 8
Const CreateNew = True

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .describe "TextStreamer class"
        Dim ts : Set ts = incl.LoadObject( "TextStreamer" )

    .it "should default to Ascii format"

        .AssertEqual ts.GetStreamFormat, Ascii

    .it "should default to create a new file"

        .AssertEqual ts.GetCreateMode, CreateNew

    .it "should default to append"

        .AssertEqual ts.GetStreamMode, ForAppending

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
