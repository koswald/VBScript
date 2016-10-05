
'Run a test or group of tests

'Usage example
''
''    With CreateObject("includer")
''        ExecuteGlobal(.read("VBSTestRunner"))
''    End With
'' 
''    With New VBSTestRunner
''        .SetSpecFolder "../spec"
''        .Run
''    End With
'
Class VBSTestRunner

    Private passing, failing, erring, foundTestFiles
    Private regex
    Private specFolder, fs, specPattern, specFile
    Private searchingSubfolders, specFileExists

    Sub Class_Initialize
        passing = 0
        failing = 0
        erring = 0
        foundTestFiles = 0
        With CreateObject("includer")
            ExecuteGlobal(.read("VBSFileSystem"))
        End With
        Set fs = New VBSFileSystem
        specFolder = ""
        SetSpecFile ""
        SetSpecPattern ".*\.spec\.vbs"
        SetSearchSubfolders False
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

    'Ensure that the specified settings are valid

    Private Sub ValidateSettings
        Dim msg

        msg = "The folder specified using SetSpecFolder must exist. A relative path is fine, relative to the calling script's folder, " & fs.SFolderName

        If Not fs.fso.FolderExists(specFolder) Then Err.Raise 1, fs.SName, msg

        msg = "Wnen SetSpecFile is used to specify a single spec file, the file specified (" & specFile & ") must exist. A relative path is fine, relative to the spec folder, " & specFolder

        If Not "" = specFile Then
            If Not fs.fso.FileExists(fs.ResolveTo(specFile, specFolder)) Then Err.Raise 1, fs.SName, msg
        End If
    End Sub

    'Method Run
    'Remark: Initiate the specified tests

    Sub Run

        ValidateSettings

        'run the test(s)

        If Len(specFile) Then
            RunTest fs.ResolveTo(specFile, specFolder)
        Else
            Set regex = New RegExp
            regex.IgnoreCase = True
            regex.Pattern = specPattern
            ProcessFiles fs.fso.GetFolder(specFolder)
        End If

        'write the result summary

        If GetErring Then
            Write_ GetErring & " erring files, "
        End If
        If GetFailing Then
            Write_ GetFailing & " failing specs, "
        End If
        If GetPassing Then
            Write_ GetPassing & " passing specs; "
        End If
        Write_ GetSpecFiles & " test files"

    End Sub 'Run

    'search for and run tests in the specified folder

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
                IncrementSpecFiles
                RunTest File.Path
            End If
        Next
    End Sub

    'run a single test file

    Private Sub RunTest(filespec)
        Dim Line
        Dim Pipe : Set Pipe = fs.sh.Exec("%ComSpec% /c cscript //nologo " & filespec)

        While Not Pipe.StdOut.AtEndOfStream
            Line = Pipe.StdOut.ReadLine
            WriteLine Line
            If "pass" = LCase(Left(Line, 4)) Then IncrementPassing
            If "fail" = LCase(Left(Line, 4)) Then IncrementFailing
        Wend

        While Not Pipe.StdErr.AtEndOfStream
            Line = Pipe.StdErr.ReadLine
            If Not "" = Line Then
                WriteLine WScript.ScriptName & ": """ & Line & """"
                IncrementErring
            End If
        Wend
    End Sub

    'Write a line to StdOut

    Private Sub WriteLine(line)
        If Not "" = line Then WScript.StdOut.WriteLine line
    End Sub

    'Write to StdOut

    Private Sub Write_(str)
        If Not "" = str Then WScript.StdOut.Write str
    End Sub

End Class 'VBSTestRunner