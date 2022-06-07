Option Explicit
Dim ea 'EncodingAnylizer object; to be tested
Dim incl 'VBScripting.Includer object
Dim baseName 'strinng
Dim suffix 'string
Dim format 'string: name of a file format

Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .describe "EncodingAnalyzer class"
        Set ea = incl.LoadObject( "EncodingAnalyzer" )

    .it "should identify an Ascii file, returning a boolean"
        baseName = "fixture/EncodingAnalyzer."
        suffix = ".txt"
        format = "Ascii"
        ea.SetFile(baseName & format & suffix)
        .AssertEqual ea.isAscii, True

    .it "should identify an Ascii file, returning a string"
        .AssertEqual ea.GetType, format

    .it "should identify a UTF16LE (Unicode) file, returning a boolean"
        format = "UTF16LE"
        ea.SetFile(baseName & format & suffix)
        .AssertEqual ea.isUTF16LE, True

    .it "should identify a UTF16LE (Unicode) file, returning a string"
        .AssertEqual ea.GetType, format

    .it "should identify a UTF16BE file, returning a boolean"
        format = "UTF16BE"
        ea.SetFile(baseName & format & suffix)
        .AssertEqual ea.isUTF16BE, True

    .it "should identify a UTF16BE file, returning a string"
        .AssertEqual ea.GetType, format

    .it "should identify a UTF7 file, returning a boolean"
        format = "UTF7"
        ea.SetFile(baseName & format & suffix)
        .AssertEqual ea.isUTF7, True

    .it "should identify a UTF7 file, returning a string"
        .AssertEqual ea.GetType, format

    .it "should identify a UTF8 file, returning a boolean"
        format = "UTF8"
        ea.SetFile(baseName & format & suffix)
        .AssertEqual ea.isUTF8, True

    .it "should identify a UTF8 file, returning a string"
        .AssertEqual ea.GetType, format

    .it "should identify a UTF32 file, returning a boolean"
        format = "UTF32"
        ea.SetFile(baseName & format & suffix)
        .AssertEqual ea.isUTF32, True

    .it "should identify a UTF32 file, returning a string"
        .AssertEqual ea.GetType, format

    .it "should get the Byte Order Mark bytes"
        .AssertEqual ea.GetByte(0) & ea.GetByte(1) & ea.GetByte(2) & ea.GetByte(3), "00254255"

End With
