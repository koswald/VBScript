' Test the VBSApp class

' The intention of the class under test is to enable the VBScript code to be identical or nearly identical whether called from an .hta or .wsf or .vbs file. Accordingly, the core test code is in a single .vbs file, .\fixture\VBSApp.fixture.vbs, that is referenced by both the .hta and .wsf test fixtures.

Option Explicit

With New VBSAppClassTester
    .RunTests
End With

Class VBSAppClassTester
    Private tester 'TestingFramework object
    Private format 'StringFormatter object
    Private sh 'WScript.Shell object
    Private fso 'Scripting.FileSystemObject
    Private outputFiles 'array of file names
    Private minusSpec, plusSpec 'defined in .config file
    Private base 'partial filespec of a fixture file

    Sub Class_Initialize
        With CreateObject( "VBScripting.Includer" )
            Execute .Read( "TestingFramework" )
            Set tester = New TestingFramework
            Execute .Read( "StringFormatter" )
            Set format = New StringFormatter
            Execute .Read("..\spec\VBSApp.spec.config")
            base = "fixture\VBSApp.fixture."
        End With
        Set sh = CreateObject( "WScript.Shell" )
        Set fso = CreateObject( "Scripting.FileSystemObject" )
        outputFiles = Array( _
            "%AppData%\VBScripting\VBSApp.htaOut.txt", _
            "%AppData%\VBScripting\VBSApp.wsfOut.txt")
        Delete(outputFiles)
    End Sub

    Sub RunTests
        tester.Describe "VBSApp class"
        RunTest "wsf", 1, "wscript.exe", 100
        RunTest "wsf", 1, "wscript.exe", 1000
        RunTest "hta", 0, "mshta.exe", 100
        RunTest "hta", 0, "mshta.exe", 1000
    End Sub

    ' Run the specified fixture file (as specified by file extension, ext), output file (index), exe, and sleep time (milliseconds), and make assertions that the outputs, as read from a file, are as expected.
    Sub RunTest(ext, index, exe, milliseconds)

        Dim stream 'text stream to read from a file
        Dim minTime, maxTime 'expected millisec
        Dim actualSleep 'actual milliseconds
        Dim j 'loop increment for timeout
        Dim instantiationMethod 'integer: loop increment
        Dim command 'string: Windows command
        Dim process 'object: sh.Exec return value
        Dim s 'string
        Dim actual, expected 'assertion arguments
        Dim iniStrings 'array: strings describing instantiation methods
        Const running = 0, finished = 1 'Exec Status
        Const COMObject = 0, NewClassName = 1 'instantiation methods

        iniStrings = Array( "COM object", "New object" )

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
                s = WScript.ScriptName & ": "
                s = s & "Excessive wait for the process"
                s = s & " '" & command & "' "
                s = s & "to finish."
                WScript.StdOut.WriteLine s
                process.Terminate
                WScript.Quit
            End If
        Loop

        ' Open the output file for reading
        Set stream = fso.OpenTextFile(Expand(outputFiles(index)))

        ' Specs and assertions

        ' The .wsf and .hta fixture files write values to .txt output files. This script reads those values from the .txt files, and the values are used in the first argument in the AssertEqual statements.
        With tester
            For instantiationMethod = ComObject To NewClassName

                'Label the test group
                s = "          ." & ext & " | " '.hta or .wsf
                s = s & iniStrings( instantiationMethod )
                s = s & " | " & milliseconds & "ms sleep"
                .ShowPendingResult
                WScript.StdOut.WriteLine s

                .It "should get a command-line argument"
                    actual = stream.ReadLine
                    expected = "arg one"
                    .AssertEqual actual, expected

                .It "should not wrap spaceless args by default"
                    actual = stream.ReadLine
                    expected = format(Array( _
                        " ""ar g ze ro"" ""arg one"" %s %s", _
                        milliseconds, ext _
                    ))
                    .AssertEqual actual, expected

                .It "should get the argument count"
                    actual = stream.ReadLine
                    expected = "4"
                    .AssertEqual actual, expected

                .It "should get the app filespec"
                    actual = stream.ReadLine
                    expected = fso.GetAbsolutePathName(base & ext)
                    .AssertEqual actual, expected

                .It "should get the app name"
                    actual = stream.ReadLine
                    expected = fso.GetFileName(base & ext)
                    .AssertEqual actual, expected

                .It "should get the base app name"
                    actual = stream.ReadLine
                    expected = fso.GetBaseName(base & ext)
                    .AssertEqual actual, expected

                .It "should get the app's filename extension"
                    actual = stream.ReadLine
                    expected = fso.GetExtensionName(base & ext)
                    .AssertEqual actual, expected

                .It "should get the app's parent folder"
                    actual = stream.ReadLine
                    expected = fso.GetParentFolderName(fso.GetAbsolutePathName(base & ext))
                    .AssertEqual actual, expected

                .It "should get the app's host .exe"
                    actual = stream.ReadLine
                    expected = exe
                    .AssertEqual actual, expected

                .It "should have a sleep method"
                    actual = stream.ReadLine
                    expected = "0"
                    .AssertEqual actual, expected

                minTime = milliseconds - minusSpec
                maxTime = milliseconds + plusSpec

                .It "should sleep for at least the min. time (" & minTime & " ms)"
                    actualSleep = stream.ReadLine * 1000
                    actual = actualSleep >= minTime
                    expected = True
                    .AssertEqual actual, expected

                .It "should sleep for at most the max. time (" & maxTime & " ms)"
                    actual = actualSleep <= maxTime
                    expected = True
                    .AssertEqual actual, expected

            Next
        End With
        stream.Close
    End Sub

    Function Expand( str )
        Expand = sh.ExpandEnvironmentStrings( str )
    End Function

    Private Sub Delete(files)
        Const Force = True
        Dim file
        For Each file In files
            file = sh.ExpandEnvironmentStrings(file)
            If fso.FileExists(file) Then
                fso.DeleteFile(file), Force
            End If
            If fso.FileExists(file) Then
                Err.Raise 51,, "Couldn't delete " & file
            End If
        Next
    End Sub

    Sub Class_Terminate
        Delete(outputFiles)
        Set fso = Nothing
        Set sh = Nothing
    End Sub
End Class
