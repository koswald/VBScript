
'Test the HTAApp class functions not
'already tested with VBSApp.spec.vbs

With CreateObject("includer")
    Execute(.read("TestingFramework"))
    Execute(.read("StringFormatter"))
    Dim outputFiles, htaFiles
    Execute(.read("..\spec\HTAApp.spec.config"))
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

    .DeleteFiles outputFiles

End With

'stream.Close
Set stream = Nothing
Set fso = Nothing
Set sh = Nothing
