
'HTAApp class

'Supports the VBSApp class, providing .hta functionality.
'
Class HTAApp
    
    'Method Sleep
    'Parameter: an integer
    'Remark: Pauses execution of the script or .hta for the specified number of milliseconds.
    Sub Sleep(ByVal milliseconds)
        milliseconds = CLng(milliseconds)
        If milliseconds - EffectiveScriptSleepOverhead > 0.0001 Then
            ScriptSleep milliseconds
        Else
            TimerSleep milliseconds
        End If
    End Sub
    
    'Private method ScriptSleep
    'Parameter: an integer
    'Remark: Private synchronous sleep method. Sleeps the specified number of milliseconds. Intended for sleeps longer than one second or so.
    Private Sub ScriptSleep(milliseconds)
        tmr.Reset
        'call the sleep script
        Dim cmd : cmd = format(Array( _
            "wscript.exe ""%s\HTAApp.sleep.vbs"" %s", _
            libraryPath, milliseconds - EffectiveScriptSleepOverhead))

        sh.Run cmd, hidden, synchronous
        'finish out with the more precise TimerSleep
        TimerSleep(milliseconds - tmr.Split * 1000)
    End Sub
    
    'Private method TimerSleep
    'Parameter: an integer
    'Remark: Private synchronous sleep method. Intended for short sleeps. Sleeps the specified number of milliseconds.
    Private Sub TimerSleep(milliseconds)
        Dim i
        tmr.Reset
        While tmr.Split * 1000 < milliseconds
            i = i + 1
        Wend
    End Sub
   
    Private oHtaObject_400BFC32009942E895C3F39EA37103DF 'must differ from calling hta's id
    Private sh, re, format
    Private filespec
    Private visible, hidden
    Private synchronous
    Private libraryPath
    Private tmr, EffectiveScriptSleepOverhead, AlwaysPrepareToSleep

    Sub Class_Initialize
        Set sh = CreateObject("WScript.Shell")
        Set re = New RegExp
        With CreateObject("includer")
            Execute(.read("HTAApp.config"))
            Execute(.read("StringFormatter"))
        End With
        Set format = New StringFormatter
        hidden = 0
        visible = 1
        synchronous = True
        filespec = Replace(Replace(Replace(document.location.href, "file:///", ""), "%20", " "), "/", "\")
        If AlwaysPrepareToSleep Then PrepareToSleep
    End Sub

    'Property ErrMsgHtaIdMissing
    'Returns a string
    'Remark: Returns the error message used when an hta:application id is required but not present.
    Property Get ErrMsgHtaIdMissing
        ErrMsgHtaIdMissing = "For command-line argument functionality, an id property must be declared in the .hta file's hta:application element."
    End Property

    'Method PrepareToSleep
    'Remark: Required before calling the Sleep method when AlwaysPrepareToSleep is False in HTAApp.config.
    Sub PrepareToSleep
        With CreateObject("includer")
            Execute(.read("VBSTimer"))
            libraryPath = .LibraryPath
        End With
        Set tmr = New VBSTimer
    End Sub

    'Property GetFilespec
    'Returns a string
    'Remark: Returns the filespec of the calling .hta file.
    Property Get GetFilespec
        GetFilespec = filespec
    End Property

    'Method SetObj
    'Parameter: the HTA id
    'Remark: Required for .hta files before accessing the arguments properties. The id is defined as a property of the .hta file's hta:application element.
    Sub SetObj(id)
        Execute("Set oHtaObject_400BFC32009942E895C3F39EA37103DF = " & id)
    End Sub

    'Property GetId
    'Returns a string
    'Remark: Extracts and returns the .hta id from the .hta file.
    'ToDo: parse the file to allow the id property to be on a separate line
    Function GetId

        'extract from the file the tag that should contain the id
        With CreateObject("includer")
            Execute(.read("VBSExtracter"))
        End With
        Dim extracter : Set extracter = New VBSExtracter
        extracter.SetFile filespec
        'extracter.SetMultiline True
        extracter.SetPattern "<hta:application[\s\S]+id\s*=[\s\S]+"
        Dim tag : tag = extracter.extract

        'extract the id from the tag
        Dim re : Set re = New RegExp
        re.IgnoreCase = True
        re.Pattern = ".+id\s*=\s*""?(\w+)""?" '? matches zero or one space or double quote; \s* matches zero or more whitespace characters
        Dim matches : Set matches = re.Execute(tag)
        On Error Resume Next
            Dim match : Set match = matches(0)
            GetId = match.Submatches(0)
            If Err Then
                If GetUserInteractive Then If vbOK = MsgBox(ErrMsgHtaIdMissing & vbLf & vbLf & filespec & vbLf & vbLf & "Do you want to open the .hta file?", vbExclamation + vbOKCancel, GetFileName) Then sh.Run "notepad """ & GetFullName & """"
                Quit
            End If
        On Error Goto 0
        'release object memory
        Set match = Nothing
        Set matches = Nothing
        Set re = Nothing
    End Function

    'Function GetArgs
    'Returns: an array
    'Remark: Returns the mshta.exe command line args as an array, including the .hta filespec, which has index 0.
    Function GetArgs
        'ensure that the hta object has been initialized
        If "HTMLGenericElement" = TypeName(oHtaObject_400BFC32009942E895C3F39EA37103DF) Then
            'expected behavior
        ElseIf "Empty" = TypeName(oHtaObject_400BFC32009942E895C3F39EA37103DF) Then
            Err.Raise 1,, ErrMsgHtaIdMissing 'hta object was not initialized
        End If
        Dim cl : cl = oHtaObject_400BFC32009942E895C3F39EA37103DF.CommandLine
        'the command line may contain two spaces between the .hta filespec and the other args. See HKCR\htafile\Shell\Open\Command
        'Convert the double space to a single space
        cl = Replace(cl, """  ", """ ")

        'surround args with double quotes,
        'if not already quoted
        cl = cl & " " 'add a trailing space, to simplify the regular expression
        re.IgnoreCase = True
        re.Global = True
        'Explanation of the regular expression
        '[^""] matches any single character except the double quote
        '+ matches one or more non-quote characters
        '? makes the match a non-greedy match
        '(?= ) non-consuming positive lookahead for a space
        re.Pattern = " ([^""]+?)(?= )" 'match only unquoted args
        cl = re.Replace(cl, " ""$1""") 'add quotes
        cl = Trim(cl) 'remove trailing space

        'convert string to array
        Dim args : args = Split(cl, """ """)
        'remove the remaining double quotes
        Dim i : For i = 0 To UBound(args)
            args(i) = Replace(args(i), """", "")
        Next
        GetArgs = args
    End Function

    Private Sub ReleaseObjectMemory
        Set sh = Nothing
        Set oHtaObject_400BFC32009942E895C3F39EA37103DF = Nothing
    End Sub

    Sub Class_Terminate
        ReleaseObjectMemory
    End Sub
End Class
