
'Pause Windows Updates to get more bandwidth. Don't forget to resume.

'For configuration settings, see the similarly named .config file in the same folder as the calling script/hta.

Class WindowsUpdatesPauser

    'Method PauseUpdates
    'Remark: Pauses Windows Updates.

    Sub PauseUpdates
        sh.Run format(Array("cmd /c netsh %s set profileparameter name=""%s"" cost=Fixed", srvcType, profileName)), hidden, synchronous
    End Sub

    'Method ResumeUpdates
    'Remark: Resumes Windows Updates.

    Sub ResumeUpdates
        sh.Run format(Array("cmd /c netsh %s set profileparameter name=""%s"" cost=Unrestricted", srvcType, profileName)), hidden, synchronous
    End Sub

    'Function GetStatus
    'Returns a string
    'Remark: Returns Metered or Unmetered. If Metered, then Windows Updates has paused to save money, incidentally not soaking up so much bandwidth. If TypeName(GetStatus) = "Empty", then the status could not be determined, possibly due to a bad network name AKA profileName.

    Function GetStatus
        Dim stat : Set stat = sh.Exec(format(Array("cmd /c netsh %s show profile name=""%s""", srvcType, profileName)))
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

    Private sh, fso, format, log 'objects
    Private thisFile
    Public appName
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
        ExecuteConfigFile defaultConfigFile, True
        'If the configFile variable was not initialized,
        'use the default path\name
        If undefined = TypeName(configFile) Then
            'we know that a custom configFile was not specified,
            'and we now alredy have the necessary configFile data,
            'so just resolve defaultConfigFile and use that
            configFile = fso.GetAbsolutePathName(defaultConfigFile)
        Else
            'a custom configFile was specified in the default configFile,
            'so get/use the data in the custom configFile
            '(except ignore configFile in the custom file)
            Dim savedConfigFile : savedConfigFile = configFile 'save
            ExecuteConfigFile configFile, False
            configFile = savedConfigFile 'restore
        End If
    End Sub

    Private Sub ExecuteConfigFile(ByVal file, FirstCall)
        file = sh.ExpandEnvironmentStrings(file)
        'if the file doesn't exist, then create it
        If Not fso.FileExists(file) Then CreateConfigFile file, FirstCall
        Dim stream : Set stream = fso.OpenTextFile(file, ForReading)
        Execute(stream.ReadAll)
        stream.Close
        Set stream = Nothing
    End Sub

    Sub CreateConfigFile(file, FirstCall)
        Dim conf : Set conf = fso.OpenTextFile(file, ForWriting, CreateNew)
        If FirstCall Then
            conf.WriteLine "'the configFile variable is optional. If it is specified, and it should be specified only in this file,"
            conf.WriteLine "'then the other variables in this file are overwritten by any variables with the same name in the specified file."
            conf.WriteLine format(Array( "''''%s = ""%s"" '%s", "configFile", "%UserProfile%\" & defaultConfigFile, "optional"))
        End If
        conf.WriteLine format(Array( "%s = ""%s"" '%s", "profileName", "My network name", "network name" ))
        conf.WriteLine format(Array( "%s = ""%s"" '%s", "srvcType", "wlan", "options: mbn (Mobile broadband), wlan (Wi-Fi)" ))
        conf.WriteLine format(Array( "%s = %s", "userInteractive", "True" ))
        conf.close
        Set conf = Nothing
    End Sub

    Sub ShowStatusError
        Dim msg : msg = format(Array( _
            "The metered status of the connection %s could not be determined. %s" & _
            "Check the profileName (network name) in %s" _
            , L2 & profileName & L2, L, L2 & configFile))
        log msg
        If Not userInteractive Then Exit Sub
        Dim msg2 : msg2 = msg & format(Array("%s Would you like %s to open the file?", L2, appName))
        If vbOK = MsgBox(msg2, vbQuestion + vbOKCancel, appName) Then
            sh.Run "notepad """ & configFile & """"
        End If
    End Sub

    Sub Class_Terminate
        Set sh = Nothing
        Set fso = Nothing
    End Sub
End Class
