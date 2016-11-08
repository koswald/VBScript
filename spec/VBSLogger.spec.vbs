
With CreateObject("includer")
    Execute(.read("VBSLogger"))
    Execute(.read("TestingFramework"))
End With

With New TestingFramework

    .describe "VBSLogger class"

        Dim log : Set log = New VBSLogger

    .it "should return the expected default log folder"

        .AssertEqual log.GetDefaultLogFolder, "%AppData%\VBScripts\logs"

    .it "should create a log folder"

        'set a custom log folder

        Dim fs : Set fs = log.fs
        Dim testLogFolder : testLogFolder = "%UserProfile%\Desktop\" & fs.fso.GetTempName
        log.SetLogFolder testLogFolder

        .AssertEqual fs.fso.FolderExists(fs.Expand(testLogFolder)), True

    .it "should expand environment variables in the log folder name"

        .AssertEqual log.GetLogFolder, fs.Expand(testLogFolder)

    .it "should create a log file name with date-stamped name"

        log.UpdateLogFilePath("September 21, 2000")

        .AssertEqual fs.fso.GetFileName(log.GetLogFilePath), "2000-09-21-Thu.txt"


    'delete the temp folder

    fs.fso.DeleteFolder(fs.Expand(testLogFolder)) 

End With
