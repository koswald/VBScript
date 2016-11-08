
'test EncodingAnalyzer.vbs

With CreateObject("includer")
    Execute(.read("EncodingAnalyzer"))
    Execute(.read("TestingFramework"))
    Execute(.read("VBSNatives"))
End With

With New TestingFramework

    .describe "EncodingAnalyzer class"

        Dim ea : Set ea = New EncodingAnalyzer

    .it "should identify an Ascii file"

        ea.SetFile("EncodingAnalyzer.sp01.txt")

        .AssertEqual ea.isAscii, True
        .AssertEqual ea.GetType, "Ascii"

    .it "should identify a UTF16LE file"

        ea.SetFile("EncodingAnalyzer.sp02.txt")

        .AssertEqual ea.isUTF16LE, True
        .AssertEqual ea.GetType, "UTF16LE"

    .it "should identify a UTF16BE file"

        ea.SetFile("EncodingAnalyzer.sp03.txt")

        .AssertEqual ea.isUTF16BE, True
        .AssertEqual ea.GetType, "UTF16BE"

    .it "should identify a UTF7 file"

        ea.SetFile("EncodingAnalyzer.sp04.txt")

        .AssertEqual ea.isUTF7, True
        .AssertEqual ea.GetType, "UTF7"

    .it "should identify a UTF8 file"

        ea.SetFile("EncodingAnalyzer.sp05.txt")

        .AssertEqual ea.isUTF8, True
        .AssertEqual ea.GetType, "UTF8"

    .it "should identify a UTF32 file"

        ea.SetFile("EncodingAnalyzer.sp06.txt")

        .AssertEqual ea.isUTF32, True
        .AssertEqual ea.GetType, "UTF32"

    .it "should get the Byte Order Mark bytes"

        .AssertEqual ea.GetByte(0) & ea.GetByte(1) & ea.GetByte(2) & ea.GetByte(3), "00254255"
End With
