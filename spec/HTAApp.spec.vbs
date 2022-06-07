'Test HTAApp class functions not
'already tested with VBSApp.spec.vbs

Option Explicit
Dim hta 'VBSHta object; to be tested
Dim errDescr 'string
Dim incl 'VBScripting.Includer object
Dim actual, expected 'variants

Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .describe "HTAApp class"
        'Note: The HTAApp class is normally instantiated within the VBSApp class
        On Error Resume Next
            Set hta = incl.LoadObject( "HTAApp" )
            errDescr = Err.Description
        On Error Goto 0
        If Not "Object required" = Left( errDescr, 15 ) Then
            Err.Raise 17,, "Unexpected error while instantiating the New HTAApp object: " & errDescr
        End If

    .it "should return a zero-element array given no args"
        actual = hta.ParseArgs("")
        expected = Array()
        .AssertEqual Join(actual, "|"), Join(expected, "|")

    .it "should raise an error if quoted str is @ right side of arg"
        On Error Resume Next
            hta.ParseArgs("""c:\some folder\some file.txt"" /f:""fg hj""")
            .AssertEqual Left(Err.Description, 36), "Invalid command-line argument syntax"
        On Error Goto 0

    .it "should raise an error if quoted str is @ left side of arg"
        On Error Resume Next
            hta.ParseArgs("""c:\some folder\some file.txt"" ""fg hj""hg""")
            .AssertEqual Left(Err.Description, 36), "Invalid command-line argument syntax"
        On Error Goto 0

    .it "should raise an error if there is an odd number of quotes"
        On Error Resume Next
            hta.ParseArgs("""gh jhyu"" """)
            .AssertEqual Left(Err.Description, 39), "There is an odd number of double quotes"
        On Error Goto 0

    .it "should return an array of arguments"
        actual = hta.ParseArgs("""C:\htaFile.hta"" ""some string with several spaces""")
        expected = Array("C:\htaFile.hta", "some string with several spaces")
        .AssertEqual Join(actual, "|"), Join(expected, "|")

    .it "should support quoted args mixed with unquoted args"
        actual = hta.ParseArgs("""C:\f o l d e r\f i l e.txt"" arg1 arg2 ""arg3"" arg4")
        expected = Array("C:\f o l d e r\f i l e.txt", "arg1", "arg2", "arg3", "arg4")
        .AssertEqual Join(actual, "|"), Join(expected, "|")

    .it "should ignore multiple spaces between arguments"
        actual = hta.ParseArgs("""C:\f o l d e r\f i l e.txt""   arg1  arg2    ""arg3""    arg4")
        expected = Array("C:\f o l d e r\f i l e.txt", "arg1", "arg2", "arg3", "arg4")
        .AssertEqual Join(actual, "|"), Join(expected, "|")

End With
