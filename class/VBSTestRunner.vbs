
'Run a test or group of tests

'Usage example
'
''   'test-launcher.vbs
''   'run this file from a console window; e.g. cscript //nologo test-launcher.vbs
'' 
''    With CreateObject("includer")
''        ExecuteGlobal(.read("VBSTestRunner"))
''    End With
'' 
''    With New VBSTestRunner
''        .SetSpecFolder "../spec" 'location of test files relative to test-launcher.vbs
''        .Run
''    End With
'
'See also TestingFramework
'
Class VBSTestRunner
    Private passing, failing, erring, foundTestFiles 'tallies
    Private regex
    Private fs, formatter, tim_r, log
    Private specFolder, specPattern, specFile 'settings
    Private searchingSubfolders
    Private runCount
    Private timeout, TimedOut
    Private TestIsFinished, TestIsRunning
    Sub Class_Initialize
        passing = 0
        failing = 0
        erring = 0
        foundTestFiles = 0
        TestIsRunning = 0
        TestIsFinished = 1
        With CreateObject("includer")
            ExecuteGlobal(.read("VBSFileSystem"))
            ExecuteGlobal(.read("StringFormatter"))
            ExecuteGlobal(.read("VBSTimer"))
            ExecuteGlobal(.read("VBSlogger"))
        End With
        Set fs = New VBSFileSystem
        Set formatter = New StringFormatter
        Set tim_r = New VBSTimer
        Set log = New VBSLogger
        specFolder = ""
        SetSpecFile ""
        SetSpecPattern ".*\.spec\.vbs"
        SetSearchSubfolders False
        SetPrecision 2
        SetRunCount 1
        SetTimeout 0
    End Sub

    Private Property Get GetPassing : GetPassing = passing : End Property
    Private Property Get GetFailing : GetFailing = failing : End Property
    Private Property Get GetErring : GetErring = erring : End Property
    Private Property Get GetSpecFiles : GetSpecFiles = foundTestFiles : End Property
    Private Sub IncrementFailing : failing = 1 + failing : End Sub
    Private Sub IncrementPassing : passing = 1 + passing : End Sub
    Private Sub IncrementErring : erring = 1 + erring : End Sub
    Private Sub IncrementSpecFiles : foundTestFiles = 1 + foundTestFiles : End Sub

    'Method SetSpecFolder
    'Parameter a folder
    'Remark: Optional. Specifies the folder containing the test files. Can be a relative path, relative to the calling script. Default is the parent folder of the calling script.

    Sub SetSpecFolder(newSpecFolder)
        specFolder = fs.Resolve(newSpecFolder)
    End Sub

    'Method SetSpecPattern
    'Parameter a regular expression
    'Remark Optional. Specifies which file types to run. Default is .*\.spec\.vbs

    Sub SetSpecPattern(newSpecPattern)
        specPattern = newSpecPattern
    End Sub

    'Method SetSpecFile
    'Parameter: a file
    'Remark Optional. Specifies a single file to test. Include the filename extension. E.g. SomeClass.spec.vbs. A relative path is OK, relative to the spec folder. If no spec file is specified, all test files matching the specified pattern will be run. See SetSpecPattern.

    Sub SetSpecFile(newSpecFile)
        specFile = newSpecFile
    End Sub

    'Method SetSearchSubfolders
    'Parameter: a boolean
    'Remark: Optional. Specifies whether to search subfolders for test files. True or False. Default is False.

    Sub SetSearchSubfolders(newSearchingSubfolders)
        searchingSubfolders = newSearchingSubfolders
    End Sub

    'Method SetPrecision
    'Parameter: 0, 1, or 2
    'Remark: Optional. Sets the number of decimal places for reporting the elapsed time. Default is 2.

    Sub SetPrecision(newPrecision) : tim_r.SetPrecision newPrecision : End Sub

    'Method SetRunCount
    'Parameter: an integer
    'Remark: Optional. Sets the number of times to run the test(s). Default is 1.

    Sub SetRunCount(newRunCount) : runCount = newRunCount : End Sub

    'Method SetTimeout
    'Parameter: an integer
    'Remark: Optional. Sets the time in seconds to wait for each test file to finish all of its specs. After this time the test file will be terminated and the other tests, if any, will be run. 0 waits indefinitely. Default is 0.

    Sub SetTimeout(newTimeout) : timeout = newTimeout : End Sub

    Private Sub ValidateSettings
        Dim msg

        msg = "The folder specified using SetSpecFolder must exist. A relative path is fine, relative to the calling script's folder, " & fs.SFolderName

        If Not fs.fso.FolderExists(specFolder) Then Err.Raise 1, fs.SName, msg

        fs.sh.CurrentDirectory = specFolder

        msg = "Wnen SetSpecFile is used to specify a single spec file, the file specified (" & specFile & ") must exist. A relative path is fine, relative to the spec folder, " & specFolder

        If Len(specFile) Then
            If Not fs.fso.FileExists(fs.ResolveTo(specFile, specFolder)) Then Err.Raise 1, fs.SName, msg
        End If
    End Sub

    'Method Run
    'Remark: Initiate the specified tests

    Sub Run

        ValidateSettings

        'run the test(s)

        Set regex = New RegExp
        regex.IgnoreCase = True
        regex.Pattern = specPattern

        Dim i : For i = 1 To runCount
            If Len(specFile) Then
                RunTest fs.ResolveTo(specFile, specFolder) 'a single test
            Else
                ProcessFiles fs.fso.GetFolder(specFolder) 'multiple tests
            End If
        Next

        'write the result summary

        If GetErring Then
            Write_ formatter.pluralize(GetErring, "erring file") & ", "
        End If
        If GetFailing Then
            Write_ formatter.pluralize(GetFailing, "failing spec") & ", "
        End If
        If GetPassing Then
            Write_ formatter.pluralize(GetPassing, "passing spec") & "; "
        End If
        Write_ formatter.pluralize(GetSpecFiles, "test file") & "; "
        Write_ "test duration: " & formatter.pluralize(tim_r, "second") & " "

    End Sub 'Run

    'run all the test files whose names match the regex pattern

    Private Sub ProcessFiles(Folder)
        Dim File, Subfolder

        If searchingSubfolders Then
            For Each Subfolder in Folder.Subfolders
                ProcessFiles Subfolder 'recurse
            Next
        End If

        For Each File In Folder.Files

            'if the file is a test/spec file, then run it

            If regex.Test(File.Name) Then
                RunTest File.Path
            End If
        Next
    End Sub

    'run a single test file

    Private Sub RunTest(filespec)
        Dim Pipe : Set Pipe = fs.sh.Exec("%ComSpec% /c cscript //nologo " & filespec)
        TimedOut = False
        Dim Line
        IncrementSpecFiles

        'wait for test to finish or time out

        If timeout > 0 Then
            WaitForTestToFinishOrTimeout(Pipe)
        End If

        'show StdOut results not already shown

        While Not Pipe.StdOut.AtEndOfStream
            WriteALineOfStdOut(Pipe)
        Wend

        'show any errors

        While Not Pipe.StdErr.AtEndOfStream
            WriteALineOfStdErr(Pipe)
        Wend

        If TimedOut Then
            Pipe.Terminate
            log fs.fso.GetBaseName(filespec) & " timed out"
        End If
    End Sub

    Private Sub WriteALineOfStdOut(Pipe)
        Dim Line
        If Not Pipe.StdOut.AtEndOfStream Then
            Line = Pipe.StdOut.ReadLine
            WriteLine Line
            If "pass" = LCase(Left(Line, 4)) Then IncrementPassing
            If "fail" = LCase(Left(Line, 4)) Then IncrementFailing
        End If
    End Sub

    Private Sub WriteALineOfStdErr(Pipe)
        Dim Line
        If Not Pipe.StdErr.AtEndOfStream Then
            Line = Pipe.StdErr.ReadLine
            If Len(Line) Then
                WriteLine WScript.ScriptName & ": """ & Line & """"
                IncrementErring
            End If
        End If
    End Sub

    Private Sub WaitForTestToFinishOrTimeout(Pipe)
        Dim startSplit : startSplit = tim_r.split
        Do
            WScript.Sleep 100 'milliseconds
            WriteALineOfStdOut(Pipe)
            If tim_r.Split - startSplit > timeout Then Exit Do
            If TestIsFinished = Pipe.status Then Exit Sub
        Loop
        TimedOut = True
    End Sub

    'Write a line to StdOut

    Private Sub WriteLine(line)
        If Len(line) Then WScript.StdOut.WriteLine line
    End Sub

    'Write to StdOut

    Private Sub Write_(str)
        If Len(str) Then WScript.StdOut.Write str
    End Sub

End Class 'VBSTestRunner
