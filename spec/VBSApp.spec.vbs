
'Test the VBSApp class

With New VBSAppClassTester
    .RunTest
End With

Class VBSAppClassTester
    Sub RunTest
        With tester
            .describe "VBSApp class"

            Setup "hta", 0
            .it "should get command-line args"
                .AssertEqual stream.ReadLine, "arg two" 'selected arg with space
            .it "should get the command-line arg string"
                .AssertEqual stream.ReadLine, " ""arg one"" ""arg two"""
            .it "should get the argument count"
                .AssertEqual stream.ReadLine, "2"
            .it "should get app filespec"
                .AssertEqual stream.ReadLine, fso.GetAbsolutePathName(base & "hta")
            .it "should get app name"
                .AssertEqual stream.ReadLine, fso.GetFileName(base & "hta")
            .it "should get the base app name"
                .AssertEqual stream.ReadLine, fso.GetBaseName(base & "hta")
            .it "should get the app's filename extension"
                .AssertEqual stream.ReadLine, fso.GetExtensionName(base & "hta")
            .it "should get the app's host .exe"
                .AssertEqual stream.ReadLine, "mshta.exe"
            .it "should have a sleep method"
                .AssertEqual stream.ReadLine, "0"
            .it "should sleep for at least the min. time"
                actualSleep = stream.ReadLine * 1000
                .AssertEqual actualSleep >= lowerLimit, True
            .it "should sleep for at most the max. time"
                .AssertEqual actualSleep <= upperLimit, True

            'Debug1

            Setup "vbs", 1
            .it "should get command-line args"
                .AssertEqual stream.ReadLine, "arg two"
            .it "should get the command-line arg string"
                .AssertEqual stream.ReadLine, " ""arg one"" ""arg two"""
            .it "should get the argument count"
                .AssertEqual stream.ReadLine, "2"
            .it "should get app filespec"
                .AssertEqual stream.ReadLine, fso.GetAbsolutePathName(base & "vbs")
            .it "should get app name"
                .AssertEqual stream.ReadLine, fso.GetFileName(base & "vbs")
            .it "should get the base app name"
                .AssertEqual stream.ReadLine, fso.GetBaseName(base & "vbs")
            .it "should get the app's filename extension"
                .AssertEqual stream.ReadLine, fso.GetExtensionName(base & "vbs")
            .it "should get the app's host .exe"
                .AssertEqual stream.ReadLine, "wscript.exe"
            .it "should have a sleep method"
                .AssertEqual stream.ReadLine, "0"
            .it "should sleep for at least the min. time"
                actualSleep = stream.ReadLine * 1000
                .AssertEqual actualSleep >= lowerLimit, True
            .it "should sleep for at most the max. time"
                .AssertEqual actualSleep <= upperLimit, True

            'Debug1

            Setup "wsf", 2
            .it "should get command-line args"
                .AssertEqual stream.ReadLine, "arg two"
            .it "should get the command-line arg string"
                .AssertEqual stream.ReadLine, " ""arg one"" ""arg two"""
            .it "should get the argument count"
                .AssertEqual stream.ReadLine, "2"
            .it "should get app filespec"
                .AssertEqual stream.ReadLine, fso.GetAbsolutePathName(base & "wsf")
            .it "should get app name"
                .AssertEqual stream.ReadLine, fso.GetFileName(base & "wsf")
            .it "should get the base app name"
                .AssertEqual stream.ReadLine, fso.GetBaseName(base & "wsf")
            .it "should get the app's filename extension"
                .AssertEqual stream.ReadLine, fso.GetExtensionName(base & "wsf")
            .it "should get the app's host .exe"
                .AssertEqual stream.ReadLine, "wscript.exe"
            .it "should have a sleep method"
                .AssertEqual stream.ReadLine, "0"
            .it "should sleep for at least the min. time"
                actualSleep = stream.ReadLine * 1000
                .AssertEqual actualSleep >= lowerLimit, True
            .it "should sleep for at most the max. time"
                .AssertEqual actualSleep <= upperLimit, True

            'Debug1

        End With
    End Sub

    Private app, tester
    Private sh, fso, stream
    Private ForReading, Synchronous, hidden
    Private base, outputFiles
    Private upperLimit, lowerLimit, testSleep, actualSleep

    Sub Class_Initialize
        With CreateObject("includer")
            Execute(.read("TestingFramework"))
            Execute(.read("..\spec\VBSApp.spec.config"))
        End With
        Set tester = New TestingFramework
        Set sh = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")
        ForReading = 1
        Synchronous = True
        hidden = 0
        base = "fixture\VBSApp.fixture."
        outputFiles = Array(base & "HtaOut.txt", _
                            base & "VbsOut.txt", _
                            base & "WsfOut.txt")
        Delete(outputFiles)
    End Sub

    'run the fixture file for the specified file extension
    Private Sub Setup(ext, index)
        tester.ShowPendingResult
        WScript.StdOut.WriteLine "          " & ext
        sh.Run "cmd /c " & base & ext & " ""arg one"" ""arg two""", hidden, Synchronous
        If Not "Empty" = TypeName(stream) Then stream.Close
        Set stream = fso.OpenTextFile(outputFiles(index))
    End Sub

    Private Sub Delete(files)
        Dim file
        For Each file In files
            If fso.FileExists(file) Then fso.DeleteFile(file)
        Next
    End Sub

    Sub Debug1
        tester.ShowPendingResult
        WScript.StdOut.WriteLine "actualSleep: " & actualSleep
        WScript.StdOut.WriteLine "upperLimit: " & upperLimit
        WScript.StdOut.WriteLine "lowerLimit: " & lowerLimit
        WScript.StdOut.WriteLine "testSleep: " & testSleep
    End Sub

    Sub Class_Terminate
        'close the text stream
        stream.Close
        'delete the output files
        Delete(outputFiles)
        'release object memory
        Set fso = Nothing
        Set sh = Nothing
    End Sub
End Class