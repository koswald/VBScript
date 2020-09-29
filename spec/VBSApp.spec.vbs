' Test the VBSApp class

' The intention of the class under test is to enable the VBScript code to be identical or nearly identical whether called from an .hta or .wsf or .vbs file. Accordingly, the core test code is in a single .vbs file that is referenced by both the .hta and .wsf test fixtures.

Option Explicit

With New VBSAppClassTester
    .RunTests
End With

Class VBSAppClassTester

    Sub RunTests
        tester.describe "VBSApp class"
        RunTest "wsf", 1, "wscript.exe", 100
        RunTest "wsf", 1, "wscript.exe", 1000
        'Debug1
        RunTest "hta", 0, "mshta.exe", 100
        RunTest "hta", 0, "mshta.exe", 1000
        'Debug1
    End Sub

    ' Run the specified fixture file (as specified by file extension), output file (index), exe, and sleep time (milliseconds). 
    Private Sub RunTest(ext, index, exe, milliseconds)

        ' Flush out result from previous test, if any.
        tester.ShowPendingResult

        ' Label the test result with the fixture file's extension name
        WScript.StdOut.WriteLine "          " & ext

        ' Run the .wsf or .hta test file. 
        command = format(Array( _
            "%s ""%s.%s"" ""ar g ze ro"" ""arg one"" %s %s", _
            exe, fso.GetAbsolutePathName(base), ext, milliseconds, ext _
        ))
        Set process = sh.Exec(command)

        ' Wait for the .wsf or .hta process to finish writing to its output file
        j = 0
        Do While Not finished = process.Status
            WScript.Sleep 50
            j = j + 1
            If j = 200 Then
                WScript.StdOut.WriteLine vbLf & "Excessive wait for the process"
                WScript.StdOut.WriteLine vbLf & command
                WScript.StdOut.WriteLine vbLf & "to finish. Please try again."
                WScript.Quit 
            End If
        Loop

        ' Open the output file for reading
        Set inStream = fso.OpenTextFile(outputFiles(index))

        ' The .wsf and .hta fixture files write values to .txt output files in the 'fixture' folder. This script reads those values from the .txt files, and the values become the first argument in the AssertEqual statements.

        ' Specifications/assertions
        With tester
            .it "should get command-line args"
                .AssertEqual inStream.ReadLine, "arg one" 'selected arg with space
            .it "should not wrap spaceless args by default"
                .AssertEqual inStream.ReadLine, format(Array( _
                    " ""ar g ze ro"" ""arg one"" %s %s", _
                    milliseconds, ext _
                ))
            .it "should get the argument count"
                .AssertEqual inStream.ReadLine, "4"
            .it "should get app filespec"
                .AssertEqual inStream.ReadLine, fso.GetAbsolutePathName(base & ext)
            .it "should get app name"
                .AssertEqual inStream.ReadLine, fso.GetFileName(base & ext)
            .it "should get the base app name"
                .AssertEqual inStream.ReadLine, fso.GetBaseName(base & ext)
            .it "should get the app's filename extension"
                .AssertEqual inStream.ReadLine, fso.GetExtensionName(base & ext)
            .it "should get the app's parent folder"
                .AssertEqual inStream.ReadLine, fso.GetParentFolderName(fso.GetAbsolutePathName(base & ext))
            .it "should get the app's host .exe"
                .AssertEqual inStream.ReadLine, exe
            .it "should have a sleep method"
                .AssertEqual inStream.ReadLine, "0"

                minTime = milliseconds - minusSpec
                maxTime = milliseconds + plusSpec
                
            .it "should sleep for at least the min. time (" & minTime & " ms)"
                actualSleep = inStream.ReadLine * 1000
                .AssertEqual actualSleep >= minTime, True
            .it "should sleep for at most the max. time (" & maxTime & " ms)"
                .AssertEqual actualSleep <= maxTime, True

            Dim minTime, maxTime
        End With

        inStream.Close
        Dim j, process, command
        Const running = 0, finished = 1
    End Sub

    Private tester, format
    Private sh, fso, inStream
    Private ForReading, Synchronous, hidden
    Private outputFiles
    Private actualSleep
    Private base, minusSpec, plusSpec 'defined in .config file

    Sub Class_Initialize
        With CreateObject("VBScripting.Includer")
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
        inStream.Close
        'delete the output files
        Delete(outputFiles)
        'release object memory
        Set fso = Nothing
        Set sh = Nothing
    End Sub
End Class
