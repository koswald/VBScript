' Integration test for the VBSLogger class and object

Option Explicit
Dim log ' The VBSLogger object under test
Dim fso ' a Scripting.FileSystemObject object
Dim incl ' a VBScripting.Includer object
Dim testLogFolder 'string: location for test files
Dim actual, expected

Set fso = CreateObject( "Scripting.FileSystemObject" )
Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .describe "VBSLogger class"
        Set log = incl.LoadObject( "VBSLogger" )

    .it "should return the expected default log folder"
        actual = log.GetDefaultLogFolder
        expected = "%AppData%\VBScripting\logs"
        .AssertEqual actual, expected

    .it "should create a log folder"
        'set a custom log folder
        testLogFolder = "%UserProfile%\Desktop\" & fso.GetTempName
        log.SetLogFolder testLogFolder
        actual = fso.FolderExists(Expand(testLogFolder))
        expected = True
        .AssertEqual actual, expected

    .it "should expand environment variables in the log folder name"
        actual = log.GetLogFolder
        expected = Expand(testLogFolder)
        .AssertEqual actual, expected

    .it "should create a log file name with date-stamped name"
        log.UpdateLogFilePath("September 21, 2000")
        actual = fso.GetFileName(log.GetLogFilePath)
        expected = "2000-09-21-Thu.txt"
        .AssertEqual actual, expected

    'cleanup

    'delete the temp folder
    fso.DeleteFolder(Expand(testLogFolder))

    'release object memory
    Set fso = Nothing

End With

Function Expand( str )
    With CreateObject( "WScript.Shell" )
        Expand = .ExpandEnvironmentStrings( str )
    End With
End Function