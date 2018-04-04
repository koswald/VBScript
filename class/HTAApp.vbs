
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

    'Return an array of command line arguments.
    'Undocumented Function is public for testability.
    Function ParseArgs(cl)
        If 0 = Len(Trim(cl)) Then ParseArgs = Array() : Exit Function
        Dim pos 'current position
        Dim char 'current character
        Dim prevChar : prevChar = " "
        Dim qCount : qCount = 0
        Dim q : q = """"
        Dim space : space = " "
        Dim argCount : argCount = 0
        Dim args : args = ""

        'read the command line, one character at a time,
        'making slight modifications
        For pos = 1 To Len(cl)
            'get the current character
            char = Mid(cl, pos, 1)
            If q = char Then qCount = qCount + 1
            If qCount mod 2 Then

                'quote count is odd...
                'validate
                If q = char And Not space = prevChar Then Err.Raise 1,, "Invalid command-line argument syntax at position " & pos & " of: " & cl
                If pos = Len(cl) Then Err.Raise 2,, "There is an odd number of double quotes in the command line arguments, " & cl

                'add the current character to the rebuild string
                args = args & char
            Else

                'quote count is even...
                'remove multiple spaces between arguments and
                'add quotes, temporarily
                'validate
                If q = prevChar And Not space = char Then Err.Raise 3,, "Invalid command-line argument syntax at position " & pos & " of: " & cl

                'rebuild arguments
                'add a leading quote to a quoteless argument
                If space = prevChar And Not space = char And Not q = char Then
                    args = args & q & char

                'add a trailing quote to a quoteless argument
                ElseIf space = char And Not space = prevChar And Not q = prevChar Then
                    args = args & q & char

                'remove multiple spaces
                ElseIf space = char And space = prevChar Then
                    'don't use this character

                'add the current character to the rebuild string
                Else
                    args = args & char
                End If
            End If
            prevChar = char
        Next

        'remove leading and trailing spaces and quotes
        args = Trim(args)
        If q = Right(args, 1) Then args = Left(args, Len(args) - 1)
        If q = Left(args, 1) Then args = Right(args, Len(args) - 1)
        ParseArgs = Split(args, """ """)
    End Function
   
    Private sh, re, format
    Private application
    Private args
    Private filespec
    Private visible, hidden
    Private synchronous
    Private libraryPath
    Private tmr, EffectiveScriptSleepOverhead, AlwaysPrepareToSleep

    Sub Class_Initialize
        Set sh = CreateObject("WScript.Shell")
        Set re = New RegExp
        With CreateObject("VBScripting.Includer")
            Execute .read("HTAApp.config")
            Execute .read("StringFormatter")
        End With
        Set format = New StringFormatter
        hidden = 0
        visible = 1
        synchronous = True
        Set application = document.getElementsByTagName("application")(0)
        args = ParseArgs(application.CommandLine)
        filespec = args(0)
        If AlwaysPrepareToSleep Then PrepareToSleep
    End Sub

    'Method PrepareToSleep
    'Remark: Required before calling the Sleep method when AlwaysPrepareToSleep is False in HTAApp.config.
    Sub PrepareToSleep
        With CreateObject("VBScripting.Includer")
            Execute .read("VBSTimer")
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

    'Function GetArgs
    'Returns: an array
    'Remark: Returns the mshta.exe command line args as an array, including the .hta filespec, which has index 0.
    Function GetArgs
        GetArgs = args
    End Function

    Private Sub ReleaseObjectMemory
        Set sh = Nothing
    End Sub

    Sub Class_Terminate
        ReleaseObjectMemory
    End Sub
End Class
