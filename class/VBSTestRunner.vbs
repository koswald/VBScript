
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

    Private syntax
    Private passing, failing, erring
    Private regex
    Private specFolder, fs, specPattern, specFile
    Private searchingSubfolders, specFileExists

    Sub Class_Initialize
        syntax = "cscript TestRunner.vbs <specfile> | -regex <pattern> [-subfolders]"
        passing = 0
        failing = 0
        erring = 0
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
    Private Sub IncrementFailing : failing = 1 + failing : End Sub
    Private Sub IncrementPassing : passing = 1 + passing : End Sub
    Private Sub IncrementErring : erring = 1 + erring : End Sub

    'Method SetSpecFolder
    'Parameter a folder
    'Remark Specifies the folder containing the test files. Can be a relative path, relative to the calling script.
    Sub SetSpecFolder(newSpecFolder)
        specFolder = fs.ResolveTo(newSpecFolder, fs.SFolderName)
    End Sub

    'Method SetSpecPattern
    'Parameter a regular expression
    'Remark Specifies which file types to run. Default is .*\.spec\.vbs
    Sub SetSpecPattern(newSpecPattern)
        specPattern = newSpecPattern
    End Sub

    'Method SetSpecFile
    'Parameter: a single file name / relative path
    'Remark Specifies a single file to test. Include the filename extension. E.g. SomeClass.spec.vbs. A relative path is OK, relative to the spec folder.
    Sub SetSpecFile(newSpecFile)
        specFile = newSpecFile
    End Sub

    'Method SetSearchSubfolders
    'Parameter: a boolean
    'Remark: Specifies whether to search subfolders for test files. True or False.
    Sub SetSearchSubfolders(newSearchingSubfolders)
        searchingSubfolders = newSearchingSubfolders
    End Sub

    Private Sub ValidateSettings
        Dim msg

        msg = "An existing spec folder must be set using SetSpecFolder. A relative path is fine, relative to the calling script's folder, " & fs.Parent(fs.SFullName)

        If Not fs.fso.FolderExists(specFolder) Then Err.Raise 1, fs.SName, msg

        msg = "The file set with SetSpecFile, " & specFile & ", must exist. A relative path is fine, relative to the spec folder, " & specFolder

        If Not "" = specFile Then
            If Not fs.fso.FileExists(fs.ResolveTo(specFile, specFolder)) Then Err.Raise 1, fs.SName, msg
        End If
    End Sub

    'Method Run
    'Remark: Initiate the specified tests
    Sub Run

        ValidateSettings

        Set regex = New RegExp
        regex.IgnoreCase = True

        'determine the test configuration and initiate the test(s)

        If Not "" = specFile Then
            'the specFile path is resolved relative to the specFolder, not the calling script's folder
            RunTest fs.ResolveTo(specFile, specFolder)
        Else
            regex.Pattern = specPattern
            ProcessFiles fs.fso.GetFolder(specFolder)
        End If

        'write result summary

        If GetErring Then
            Write_ GetErring & " erring, "
        End If
        If GetFailing Then
            Write_ GetFailing & " failing, "
        End If
        Write_ GetPassing & " passing"

    End Sub 'Run

    'use the regex filter to process multiple spec files

    Private Sub ProcessFiles(Folder)
        Dim File, Subfolder

        If searchingSubfolders Then
            For Each Subfolder in Folder.Subfolders
                ProcessFiles Subfolder 'recurse
            Next
        End If

        For Each File In Folder.Files

            'if the file is a test/spec file, then run it

            If regex.Test(File.Name) Then RunTest File.Path
        Next
    End Sub

    'run a single spec file

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