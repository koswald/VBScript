
'Open a file as a text stream for reading, writing, or appending.

'<h5> Methods for use with the text stream that is returned by the Open method: </h5>

'<p> <em> Reading methods: </em> Read, ReadLine, ReadAll <br /> <em> Writing methods: </em> Write, WriteLine, WriteBlankLines <br /> <em> Reading or Writing methods: </em> Close, Skip, SkipLine <br /> <em> Reading or writing properties: </em> AtEndOfLine, AtEndOfStream, Column, Line </p>

Class TextStreamer

    Private file, folder, fileName
    Private StreamMode, AllowToCreateNew, StreamFormat, viewer, viewerProcess
    Private oStreamConstants, oVBSFileSystem

    Sub Class_Initialize 'event fires on object instantiation
        With CreateObject("includer") 'get class dependencies
            Execute(.read("VBSFileSystem"))
            Execute(.read("StreamConstants"))
        End With

        Set oVBSFileSystem = New VBSFileSystem
        Set oStreamConstants = New StreamConstants

        SetForAppending 'set defaults for Private members
        SetCreateNew
        SetAscii
        SetViewer "Notepad"
        SetFolder "%UserProfile%\Desktop"
        SetFileName fso.GetBaseName(fso.GetTempName) & ".txt"
    End Sub

    Property Get fs : Set fs = oVBSFileSystem : End Property
    Property Get sc : Set sc = oStreamConstants : End Property

    Property Get n : Set n = fs.n : End Property

    Property Get shell : Set shell = fs.sh : End Property
    Property Get sh : Set sh = fs.sh : End Property

    Property Get fso : Set fso = fs.fso : End Property

    Property Get args : Set args = a : End Property
    Property Get a : Set a = fs.a : End Property

    'Property Open
    'Returns an object
    'Remark: Returns a text stream object according to the specified settings (methods beginning with Set...)
    Property Get Open
        Set Open = fso.OpenTextFile(fs.Expand(file), StreamMode, AllowToCreateNew, StreamFormat)
    End Property

    'Method SetFile
    'Parameter: a filespec
    'Remark: Specifies the file to be opened by the text streamer. Can include environment variable names. The default file is a random-named .txt file on the desktop.
    Sub SetFile(newFile)
        file = newFile
        folder = fso.GetParentFolderName(file)
        fileName = fso.GetFileName(file)
    End Sub

    'Method SetFolder
    'Parameter: a folder
    'Remark: Specifies the folder of the file to be opened by the text streamer. Can include environment variables. Default is %UserProfile%\Desktop
    Sub SetFolder(newFolder)
        folder = newFolder
        If Not fso.FolderExists(folder) Then fs.MakeFolder folder
        SetFile folder & "\" & fileName
    End Sub

    'Method SetFileName
    'Parameter: a file name
    'Remark: Specifies the file name, including extension, of the file to be opened by the text streamer. Default is a randomly named .txt file.
    Sub SetFileName(newFileName)
        fileName = newFileName
        SetFile folder & "\" & fileName
    End Sub

    'Method SetForReading
    'Remark Prepares the text stream to be opened for reading
    Sub SetForReading : StreamMode = 1 : End Sub

    'Method SetForWriting
    'Remark Prepares the text stream to be opened for writing
    Sub SetForWriting : StreamMode = 2 : End Sub

    'Method SetForAppending
    'Remark Prepares the text stream to be opened for appending (default)
    Sub SetForAppending : StreamMode = 8 : End Sub

    'Method SetCreateNew
    'Remark: Allows a new file to be created (default)
    Sub SetCreateNew : AllowToCreateNew = True : End Sub

    'Method SetDontCreateNew
    'Remark Prevents a new file from being created if the file doesn't already exist
    Sub SetDontCreateNew : AllowToCreateNew = False : End Sub

    'Method SetAscii
    'Remark: Sets the expectation that the file will be Ascii (default)
    Sub SetAscii : StreamFormat = 0 : End Sub

    'Method SetUnicode
    'Remark: Sets the expectation that the file will be Unicode
    Sub SetUnicode : StreamFormat = -1 : End Sub

    'Method SetSystemDefault
    'Remark: Uses Ascii or Unicode according to the system default
    Sub SetSystemDefault : StreamFormat = -2 : End Sub

    'Method View
    'Remark: Opens the file for viewing
    Sub View : Set viewerProcess = sh.Exec("""" & viewer & """ """ & file & """") : End Sub

    'Method CloseViewer
    'Remark: Close the file viewer. From the docs: Use the Terminate method only as a last resort since some applications do not clean up properly. As a general rule, let the process run its course and end on its own. The Terminate method attempts to end a process using the WM_CLOSE message. If that does not work, it kills the process immediately without going through the normal shutdown procedure.

    Sub CloseViewer : viewerProcess.Terminate : End Sub

    'Method SetViewer
    'Parameter: filespec
    'Remark: Sets the filespec of an alternate file viewer to use with the View method.The default viewer is Notepad.
    Sub SetViewer(pViewer) : viewer = pViewer : End Sub

    'Method Delete
    'Remark: Deletes the streamer file
    Sub Delete
        'WScript.Sleep 10 'give time for the file to open in Notepad before deleting
        fso.DeleteFile(fs.Expand(file))
    End Sub

    'Method: Run
    'Remark: Open/Run the file, assuming it has an executable file extension.
    Sub Run : sh.Run """" & file & """" : End Sub

    'Property GetFile
    'Returns a filespec
    'Remark: Returns the filespec of the file that is open or set to be opened by the text streamer. Environment variables are not expanded.
    Property Get GetFile : GetFile = file : End Property

    'Property GetFileName
    'Returns a file name
    'Remark: Returns the file name of the file that is open or set to be opened by the text streamer. Environment variables are not expanded.
    Property Get GetFileName : GetFileName = fileName : End Property

    'Property GetFolder
    'Returns a folder
    'Remark: Returns the folder of the file that is open or set to be opened by the text streamer. Environment variables are not expanded.
    Property Get GetFolder : GetFolder = folder : End Property

    'Property GetCreateMode
    'Returns a boolean
    'Remark: Gets the current CreateMode setting. Returns one of these stream constants: bDontCreateNew or bCreateNew.
    Property Get GetCreateMode : GetCreateMode = AllowToCreateNew : End Property

    'Property GetStreamMode
    'Returns an integer
    'Remark: Gets the current StreamMode setting. Returns one of these stream constants: iForReading, iForWriting, iForAppending
    Property Get GetStreamMode : GetStreamMode = StreamMode : End Property

    'Property GetStreamFormat
    'Returns a tristate boolean
    'Remark: Gets the current StreamFormat setting. Returns one of these stream constants: tbAscii, tbUnicode, tbSystemDefault
    Property Get GetStreamFormat : GetStreamFormat = StreamFormat : End Property

    Sub Class_Terminate
        Set viewerProcess = Nothing
    End Sub

End Class
