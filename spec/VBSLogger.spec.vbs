
With CreateObject("includer")
    Execute(.read("VBSLogger"))
    Execute(.read("TestingFramework"))
End With

Dim log : Set log = New VBSLogger
Dim fs : Set fs = log.fs

With New TestingFramework

    .describe "VBSLogger class"

    .it "should return the expected default log folder"

        .AssertEqual log.GetDefaultLogFolder, "%AppData%\VBScripts\logs"

    .it "should create a log folder"

        'set the path create the log folder with a unique name

        Dim tempName : tempName = fs.fso.GetTempName
        Dim testLogFolder : testLogFolder = "%UserProfile%\Desktop\" & tempName
        log.SetLogFolder testLogFolder 'set the path and create the folder

        .AssertEqual True, fs.fso.FolderExists(fs.Expand(testLogFolder))

    .it "should expand environment variables in the log folder name"

        .AssertEqual log.GetLogFolder, fs.Expand(testLogFolder)


    .it "should create a log file name with date-stamped name"

        Dim testDate : testDate = "September 21, 2000"

        log.UpdateLogFilePath(testDate)
        Dim expectedFileName : expectedFileName = "2000-09-21-Thu.txt"

        .AssertEqual expectedFileName, fs.fso.GetFileName(log.GetLogFilePath)

    fs.fso.DeleteFolder(fs.Expand(testLogFolder)) 'delete the temp folder

End With
