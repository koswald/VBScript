
'Test the VBSApp class

With New VBSAppClassTester
    .RunTests
End With

Class VBSAppClassTester

    'Method RunTests
    'Remark: Run a series of tests. 
    'The intention of the class under test is
    'for the VBScript code to be identical or nearly
    'identical whether called from an .hta or .vbs file
    Sub RunTests
        tester.describe "VBSApp class"
        RunTest "wsf", 1, "wscript.exe", 100
        RunTest "wsf", 1, "wscript.exe", 1000
        'Debug1
        RunTest "hta", 0, "mshta.exe", 100
        RunTest "hta", 0, "mshta.exe", 1000
        'Debug1
    End Sub

    'run the specified fixture file (specified by file extension), output file (index), exe, and time (milliseconds)
    Private Sub RunTest(ext, index, exe, milliseconds)
        tester.ShowPendingResult
        WScript.StdOut.WriteLine "          " & ext
        sh.Run format(Array( _
            "cmd /c %s%s ""ar g ze ro"" ""arg one"" %s %s", _
            base, ext, milliseconds, ext _
        )), hidden, Synchronous
        If Not "Empty" = TypeName(stream) Then stream.Close 'close the previous test's text stream, unless this is the first test
        Set stream = fso.OpenTextFile(outputFiles(index))
        With tester
            .it "should get command-line args"
                .AssertEqual stream.ReadLine, "arg one" 'selected arg with space
            .it "should not wrap spaceless args by default"
                .AssertEqual stream.ReadLine, format(Array(" ""ar g ze ro"" ""arg one"" %s %s", milliseconds, ext))
            .it "should get the argument count"
                .AssertEqual stream.ReadLine, "4"
            .it "should get app filespec"
                .AssertEqual stream.ReadLine, fso.GetAbsolutePathName(base & ext)
            .it "should get app name"
                .AssertEqual stream.ReadLine, fso.GetFileName(base & ext)
            .it "should get the base app name"
                .AssertEqual stream.ReadLine, fso.GetBaseName(base & ext)
            .it "should get the app's filename extension"
                .AssertEqual stream.ReadLine, fso.GetExtensionName(base & ext)
            .it "should get the app's parent folder"
                .AssertEqual stream.ReadLine, fso.GetParentFolderName(fso.GetAbsolutePathName(base & ext))
            .it "should get the app's host .exe"
                .AssertEqual stream.ReadLine, exe
            .it "should have a sleep method"
                .AssertEqual stream.ReadLine, "0"

                Dim minTime, maxTime
                minTime = milliseconds - minusSpec
                maxTime = milliseconds + plusSpec
                
            .it "should sleep for at least the min. time (" & minTime & " ms)"
                actualSleep = stream.ReadLine * 1000
                .AssertEqual actualSleep >= minTime, True
            .it "should sleep for at most the max. time (" & maxTime & " ms)"
                .AssertEqual actualSleep <= maxTime, True
        End With
    End Sub

    Private app, tester, format
    Private sh, fso, stream
    Private ForReading, Synchronous, hidden
    Private outputFiles
    Private actualSleep
    Private base, minusSpec, plusSpec 'defined in .config file

    Sub Class_Initialize
        With CreateObject("includer")
            Execute .read("TestingFramework")
            Execute .read("StringFormatter")
            Execute .read("..\spec\VBSApp.spec.config")
        End With
        Set tester = New TestingFramework
        Set format = New StringFormatter
        Set sh = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")
        ForReading = 1
        Synchronous = True
        hidden = 0
        outputFiles = Array( _
            base & "htaOut.txt", _
            base & "wsfOut.txt")
        Delete(outputFiles)
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
        WScript.StdOut.WriteLine "plusSpec: " & plusSpec
        WScript.StdOut.WriteLine "minusSpec: " & minusSpec
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
