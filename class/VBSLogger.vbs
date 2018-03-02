
'A lightweight VBScript logger

'Instantiation 
'<pre>     With CreateObject("VBScripting.Includer") <br />         Execute .read("VBSLogger") <br />     End With <br />     Dim log : Set log = New VBSLogger </pre>
'
'Usage method one. This method has the advantage that the log doesn't remain open, allowing other scripts to write to the log.
' <pre>     log "test one" </pre>

'Usage method two. This method has the advantage that the name of the calling script is not written on each line of the log.
' <pre>     log.Open <br />     log.Write "test two" <br />     log.Close </pre>
'
Class VBSLogger 'Logger for use in VBScript files

'    Private oTimeFunctions, oTextStreamer
    Private streamer, dt, fs
    Private sh, fso
    Private stream, logFile, logFolder, viewer
    Private scriptName, scriptFullName

    Sub Class_Initialize
        WIth CreateObject("VBScripting.Includer") 'get class dependencies
            Execute .read("TimeFunctions")
            Execute .read("TextStreamer")
            Execute .read("VBSFileSystem")
        End With
        Set dt = New TimeFunctions
        Set streamer = New TextStreamer
        Set fs = New VBSFileSystem

        Set sh = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")

        On Error Resume Next
            scriptFullName = WScript.ScriptFullName
            If Err Then scriptFullName = Replace(Replace(document.location.href, "file:///", ""), "%20", "") '.hta file
        On Error Goto 0
        scriptName = fso.GetFileName(scriptFullName)
        SetLogFolder(GetDefaultLogFolder)
        SetViewer(Notepad)
        dt.LetDOWBeAbbreviated = True 'DOW = day of the week
        UpdateLogFilePath(Now)
    End Sub

    Property Get Notepad : Notepad = "Notepad" : End Property

    'Method Log
    'Parameter: a string
    'Remark: Opens the log file, writes the specified string, then closes the log file. This is the default method for the VBSLogger class.
    Public Default Sub Log(msg) 'open the log file for appending, write the message, and then close the text stream
        PrivateOpen
        stream.WriteLine(dt.GetFormattedTime(Now) & " - " & scriptName & " - " & msg)
        PrivateClose
    End Sub

    'Method SetLogFolder
    'Parameter: a folder path
    'Remark: Optional. Customize the log folder. The folder will be created if it does not exist. Environment variables are allowed. See GetDefaultLogFolder.
    Sub SetLogFolder(newLogFolder) 'set the log folder; create it if necesssary
        logFolder = fs.Expand(newLogFolder)
        If Not fs.MakeFolder(logFolder) Then Err.Raise 1, "VBSLogger.SetLogFolder", "Failed to create log folder " & logFolder
    End Sub

    Property Get GetLogFolder : GetLogFolder = logFolder : End Property

    Sub UpdateLogFilePath(date_) 'ensure that the filespec for the log file reflects the specified/current date
        logFile = logFolder & "\" & dt.GetFormattedDay(date_) & ".txt"
    End Sub

    Private Sub PrivateOpen 'open the log file as a text stream for appending
        UpdateLogFilePath(Now)
        streamer.SetFile(logFile)
        PrivateClose 'attempt to close the stream, in case it is already open
        Set stream = streamer.Open
    End Sub

    'Method Open
    'Remark: Opens the log file for writing. The log file is opened and remains open for writing. While it is open, other processes/scripts will be unable to write to it.
    Sub Open
        PrivateOpen
        stream.WriteLine(dt.GetFormattedTime(Now) & " - log opened by " & scriptName)
    End Sub

    'Method Write
    'Parameter: a string
    'Remark: Writes the specified string to the log file.
    Sub Write(msg) 'write to the log with timestamp
        stream.WriteLine(dt.GetFormattedTime(Now) & " - " & msg)
    End Sub

    'Method Close
    'Remark: Closes the log file text stream, enabling other process to write to it.
    Sub Close
        stream.WriteLine(dt.GetFormattedTime(Now) & " - log closed by " & scriptName)
        PrivateClose
    End Sub

    Private Sub PrivateClose 'close the text stream and free up object memory
        On Error Resume Next
            stream.Close
            Set stream = Nothing
        On Error Goto 0
    End Sub

    'Method View
    'Remark: Opens the log file for viewing. Notepad is the default editor. See SetViewer.
    Sub View 'open the log file for viewing in a text editor
        If fso.FileExists(GetLogFilePath) Then
            sh.Run """" & viewer & """ """ & logFile & """"
        Else
            Dim msg : msg = "Today's log file hasn't been created " & _
                "yet. Do you want to open the log folder?"
            If vbOK = MsgBox(msg, vbOKCancel + vbQuestion) Then
                ViewFolder
            End If
        End If
    End Sub

    'Method SetViewer
    'Parameter: a filespec
    'Remark: Optional. Customize the program that the View method uses to view log files. Default: Notepad.
    Sub SetViewer(newViewer) : viewer = newViewer : End Sub

    'Method ViewFolder
    'Remark: Open the log folder
    Sub ViewFolder 'open Windows File Explorer at the log folder
        sh.Run "explorer """ & logFolder & """"
    End Sub

    'Property WordPad
    'Returns a filespec
    'Remark: Can be used as the argument for the SetViewer method in order to open files with WordPad when the View method is called.
    Property Get WordPad : Wordpad = "%ProgramFiles%\Windows NT\Accessories\wordpad.exe" : End Property

    'Property GetDefaultLogFolder
    'Returns a folder
    'Remark: Retrieves the default log folder, %AppData%\VBScripts\logs
    Property Get GetDefaultLogFolder
        GetDefaultLogFolder                 = "%AppData%\VBScripts\logs"
    End Property

    'Property GetLogFilePath
    'Returns a filespec
    'Remark: Retreives the filespec for the log file, with environment variables expanded. Default: &lt;GetDefaultLogFolder&gt;\YYYY-MM-DD-DayOfWeek.txt
    Property Get GetLogFilePath : GetLogFilePath = logFile : End Property

    Sub Class_Terminate 'event fires when the logger instance goes out of scope or is Set to Nothing
        PrivateClose
    End Sub

End Class
