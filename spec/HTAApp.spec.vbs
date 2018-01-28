
'Test HTAApp class functions not
'already tested with VBSApp.spec.vbs

With CreateObject("Includer")
    Execute .Read("TestingFramework")
    Execute .Read("StringFormatter")
    Dim outputFiles, htaFiles
    Execute .Read("..\spec\HTAApp.spec.config")
    Execute .Read("HTAApp")
    On Error Resume Next
        Dim hta : Set hta = New HTAApp
        Dim errDescr : errDescr = Err.Description
    On Error Goto 0
    If Not "Object required" = errDescr Then Err.Raise 1,, "Unexpected error while instantiating the New HTAApp object: " & errDescr
End With
Dim format : Set format = New StringFormatter
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Const ForReading = 1
Const hidden = 0
Const synchronous = True
Dim stream
Dim sh : Set sh = CreateObject("WScript.Shell")
With New TestingFramework

    .describe "HTAApp class"

    .it "should raise an err if command-line args are used without an .hta id"
        sh.Run format(Array( _
            "cmd /c mshta ""%s"" arg_0 arg_1", _
            fso.GetAbsolutePathName(htaFiles(iHta_NoIdHasArgs)) _
        )), hidden, synchronous
        Set stream = fso.OpenTextFile(outputFiles(iOutput_NoIdHasArgs), ForReading)
        .AssertEqual stream.ReadLine, "For command-line argument functionality, an id property must be declared in the .hta file's hta:application element."
        stream.Close

    .it "should not require an id if command-line args aren't used"
        sh.Run format(Array( _
            "cmd /c mshta ""%s""", _
            fso.GetAbsolutePathName(htaFiles(iHta_NoIdNoArgs)) _
        )), hidden, synchronous
        Set stream = fso.OpenTextFile(outputFiles(iOutput_NoIdNoArgs), ForReading)
        .AssertEqual stream.ReadLine, ""
        stream.Close

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

    .it "should remove multiple spaces between arguments"
        actual = hta.ParseArgs("""C:\f o l d e r\f i l e.txt""   arg1  arg2    ""arg3""    arg4")
        expected = Array("C:\f o l d e r\f i l e.txt", "arg1", "arg2", "arg3", "arg4")
        .AssertEqual Join(actual, "|"), Join(expected, "|")

    .DeleteFiles outputFiles

End With

'stream.Close
Set stream = Nothing
Set fso = Nothing
Set sh = Nothing
