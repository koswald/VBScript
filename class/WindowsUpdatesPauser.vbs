
'Pause Windows Updates to get more bandwidth. Don't forget to resume.

'For configuration settings, see the similarly named .config file in the same folder as the calling script/hta.

'You can use a configFile variable in the first .config file pointing to a second .config file in the location of your choosing, using environment variables, if desired.

Class WindowsUpdatesPauser

    'Method PauseUpdates
    'Remark: Pauses Windows Updates.

    Sub PauseUpdates
        sh.Run format(Array( _
            "cmd /c netsh %s set profileparameter name=""%s"" cost=Fixed", _
            srvcType, profileName)), hidden, synchronous
    End Sub

    'Method ResumeUpdates
    'Remark: Resumes Windows Updates.

    Sub ResumeUpdates
        sh.Run format(Array( _
            "cmd /c netsh %s set profileparameter name=""%s"" cost=Unrestricted", _
            srvcType, profileName)), hidden, synchronous
    End Sub

    'Function GetStatus
    'Returns a string
    'Remark: Returns Metered or Unmetered. If Metered, then Windows Updates has paused to save money, incidentally not soaking up so much bandwidth. If TypeName(GetStatus) = "Empty", then the status could not be determined, possibly due to a bad network name AKA profileName.

    Function GetStatus
        Dim stat : Set stat = sh.Exec(format(Array( _
            "cmd /c netsh %s show profile name=""%s""", _
            srvcType, profileName)))
        Dim line
        While Not stat.StdOut.AtEndOfStream
            'get the status by parsing the output of the show profile command
            line = stat.StdOut.ReadLine
            If CBool(InStr(line, "Unrestricted")) And CBool(InStr(line, "Cost")) Then
                GetStatus = unmetered
                Exit Function
            ElseIf ( CBool(InStr(line, "Fixed")) Or CBool(InStr(line, "Variable")) ) And CBool(InStr(line, "Cost")) Then
                GetStatus = metered
                Exit Function
            End If
        Wend
        ShowStatusError
    End Function

    'Function GetAppName
    'Returns a string
    'Remark: Returns the base name of the calling script

    Function GetAppName : GetAppName = appName : End Function

    Private sh, fso, format, log 'objects
    Private thisFile
    Private appName
    Private profileName, srvcType, userInteractive, configFile 'data kept in .config file
    Private defaultConfigFile
    Private metered, unmetered, undefined ' "enums"
    Private hidden, synchronous 'Run params #2 & #3
    Private ForReading, ForWriting 'OpenTextFile param #2
    Private CreateNew 'OpenTextFile param #3
    Private L, L2

    Sub Class_Initialize
        Set sh = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")
        userInteractive = True 'defaults
        unmetered = "Unmetered" ' "enums"
        metered = "Metered"
        undefined = "Empty"
        hidden = 0 'constants
        synchronous = True
        ForReading = 1
        ForWriting = 2
        CreateNew = True
        L = vbLf 'other
        L2 = L & L
        Dim reader : Set reader = CreateObject("includer")
        Execute(reader.read("StringFormatter"))
        Execute(reader.read("VBSLogger"))
        Set format = New StringFormatter
        Set log = New VBSLogger
        Set reader = Nothing
        On Error Resume Next
            thisFile = Replace(Replace(document.location.href, "file:///", ""), "%20", " ") 'called by .hta
            If Err Then thisFile = WScript.ScriptFullName 'called by script
        On Error Goto 0
        appName = fso.GetBaseName(thisFile)
        defaultConfigFile = appName & ".config"
        ReadConfigFiles
   End Sub

    'Sub ReadConfigFiles
    'Remark: Read profileName, srvcType, userInteractive and, possibly, configFile from .config file(s)

    Private Sub ReadConfigFiles
        'make the current directory the parent folder of this file,
        'so that the first place to look for a .config file doesn't
        'depend on a shortcut's working directory
        sh.CurrentDirectory = fso.GetParentFolderName(thisFile)
        'read the default (first) .config file
        ExecuteConfigFile defaultConfigFile, True 'True => first .config file
        'determine whether the configFile variable was initialized
        If undefined = TypeName(configFile) Then
            'a second configFile was not specified,
            'and we should now have the necessary configFile data,
            'so just resolve defaultConfigFile and use that.
            configFile = fso.GetAbsolutePathName(defaultConfigFile)
        Else
            'a second configFile was specified in the default configFile,
            'so get/use the data in the second configFile
            '(except ignore configFile if present in the second file)
            Dim savedConfigFile : savedConfigFile = configFile 'save
            'read the second .config file
            ExecuteConfigFile configFile, False 'False => second .config file
            configFile = savedConfigFile 'restore
        End If
    End Sub

    'Read the .config file
    'To do: parse the file rather than just execute it!

    Private Sub ExecuteConfigFile(ByVal file, FirstCall)
        file = sh.ExpandEnvironmentStrings(file)
        'if the file doesn't exist, then create it
        If Not fso.FileExists(file) Then CreateConfigFile file, FirstCall
        Dim stream : Set stream = fso.OpenTextFile(file, ForReading)
        Execute(stream.ReadAll)
        stream.Close
        Set stream = Nothing
    End Sub

    'Create a new .config file

    Private Sub CreateConfigFile(file, FirstCall)
        Dim conf : Set conf = fso.OpenTextFile(file, ForWriting, CreateNew)
        If FirstCall Then
            'comments pertaining only to the first .config file,
            'the one in the same folder as the calling script
            conf.WriteLine "'WARNING:"
            conf.WriteLine "'If the configFile variable is specified in this file,"
            conf.WriteLine "'then any other variables in this file are ignored and"
            conf.WriteLine "'may be safely deleted."
        Else
            'comments pertaining only to the second .config file, if any
            conf.WriteLine "'WARNING:"
            conf.WriteLine "'if a configFile variable is specified in this file,"
            conf.WriteLine "'then it will be ignored"
        End If
        'For each variable, write variableName = defaultValue ['comment]
        conf.WriteLine format(Array( "%s = ""%s"" '%s", "profileName", "My network name", "network name" ))
        conf.WriteLine format(Array( "%s = ""%s"" '%s", "srvcType", "wlan", "options: mbn (Mobile broadband), wlan (Wi-Fi)" ))
        conf.WriteLine format(Array( "%s = %s", "userInteractive", "True" ))
        conf.close
        Set conf = Nothing
    End Sub

    'Log a status error, and if userInteractive, display it

    Private Sub ShowStatusError
        Dim msg : msg = format(Array( _
            "The metered status of the connection %s could not be determined. %s" & _
            "Check the profileName (network name) in %s", _
            L2 & profileName & L2, L, L2 & configFile))
        log msg
        If Not userInteractive Then Exit Sub
        Dim msg2 : msg2 = msg & format(Array( _
            "%s Would you like %s to open the file?", L2, appName))
        If vbOK = MsgBox(msg2, vbQuestion + vbOKCancel, appName) Then
            sh.Run "notepad """ & configFile & """"
        End If
    End Sub

    Sub Class_Terminate
        Set sh = Nothing
        Set fso = Nothing
        format.Class_Terminate
        log.Class_Terminate
    End Sub
End Class
