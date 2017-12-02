
With CreateObject("includer")
    Execute .read("VBSLogger")
    Execute .read("TestingFramework")
    Execute .read("VBSFileSystem")
End With
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim fs : Set fs = New VBSFileSystem

With New TestingFramework

    .describe "VBSLogger class"
        Dim log : Set log = New VBSLogger

    .it "should return the expected default log folder"
        .AssertEqual log.GetDefaultLogFolder, "%AppData%\VBScripts\logs"

    .it "should create a log folder"
        'set a custom log folder
        Dim testLogFolder : testLogFolder = "%UserProfile%\Desktop\" & fso.GetTempName
        log.SetLogFolder testLogFolder
        .AssertEqual fso.FolderExists(fs.Expand(testLogFolder)), True

    .it "should expand environment variables in the log folder name"
        .AssertEqual log.GetLogFolder, fs.Expand(testLogFolder)

    .it "should create a log file name with date-stamped name"
        log.UpdateLogFilePath("September 21, 2000")
        .AssertEqual fso.GetFileName(log.GetLogFilePath), "2000-09-21-Thu.txt"

    'cleanup

    'delete the temp folder
    fso.DeleteFolder(fs.Expand(testLogFolder))

    'release object memory
    Set fso = Nothing

End With
