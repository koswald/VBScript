
'Pause Windows Updates to get more bandwidth. Don't forget to resume.

'For configuration settings, see the .config file in %AppData%\VBScripting that has the same base name as the calling script/hta.
'
Class WindowsUpdatesPauser

    'Method PauseUpdates
    'Remark: Pauses Windows Updates.
    Sub PauseUpdates
        sh.Run format(Array( _
            "cmd /c netsh %s set profileparameter name=""%s"" cost=Fixed", _
            srvcType, profileName)), hidden, synchronous
        GetStatus
    End Sub

    'Method ResumeUpdates
    'Remark: Resumes Windows Updates.
    Sub ResumeUpdates
        sh.Run format(Array( _
            "cmd /c netsh %s set profileparameter name=""%s"" cost=Unrestricted", _
            srvcType, profileName)), hidden, synchronous
        GetStatus
    End Sub

    'Function GetStatus
    'Returns a string
    'Remark: Returns Metered or Unmetered. If Metered, then Windows Updates has paused to save money, incidentally not soaking up so much bandwidth. If TypeName(GetStatus) = "Empty", then the status could not be determined, possibly due to a bad network name (internal name: profileName).
    Function GetStatus
        Dim stat : Set stat = sh.Exec(format(Array( _
            "cmd /c netsh %s show profile name=""%s""", _
            srvcType, profileName)))
        Dim line
        Dim currentProfile : currentProfile = False
        While Not stat.StdOut.AtEndOfStream
            'get the status by parsing the output of the show profile command
            line = stat.StdOut.ReadLine
            If InStr(line, profileName) Then currentProfile = True
            If Not currentProfile Then
                'disregard other profiles; i.e. disregard the profile having the same name except with all upper case letters, if that name is not exactly the same as profileName;
                'this is necessary because the show profile command may show two sets of results: if it does, then the first set, to be disregarded, is for a similarly named but distinct profile but with all upper case letters in the name.
            ElseIf CBool(InStr(line, "Unrestricted")) And CBool(InStr(line, "Cost")) Then
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
    Function GetAppName : GetAppName = app.GetBaseName : End Function

    'Function GetProfileName
    'Returns a string
    'Remark: Returns the name of the network. The name is set by editing the .config file in %AppData%\VBScripting that has the same base name as the calling script/hta.
    Property Get GetProfileName : GetProfileName = profileName : End Property

    'Function GetServiceType
    'Returns a string
    'Remark: Returns the service type
    Function GetServiceType : GetServiceType = srvcType : End Function

    'Method OpenConfigFile
    'Remark: Opens the .config file for editing.
    Sub OpenConfigFile : sh.Run "notepad """ & configFile & """" : End Sub

    'Private method ReadConfigFile
    'Remark: Read profileName, srvcType, userInteractive from the .config file
    Private Sub ReadConfigFile
        Dim file : file = sh.ExpandEnvironmentStrings(configFile)
        'if the file doesn't exist, then create it
        If Not fso.FileExists(file) Then CreateConfigFile file
        Dim stream : Set stream = fso.OpenTextFile(file, ForReading)
        Execute(stream.ReadAll)
        stream.Close
        Set stream = Nothing
    End Sub

    'Create a new .config file
    Private Sub CreateConfigFile(file)
        Dim conf : Set conf = fso.OpenTextFile(file, ForWriting, CreateNew)
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
            "The metered status of the connection %scould not be determined. %s" & _
            "Check the profileName (network name) in %s", _
            L2 & profileName & L2, L, L & configFile))
        log msg
        If Not userInteractive Then Exit Sub
        Dim msg2 : msg2 = msg & format(Array( _
            "%s Would you like %s to open the file?", L2, app.GetBaseName))
        If vbOK = MsgBox(msg2, vbInformation + vbOKCancel, app.GetBaseName) Then
            OpenConfigFile
        End If
    End Sub

    Private sh, fso, format, log, app 'objects
    Private profileName, srvcType, userInteractive, configFile 'data kept in .config file
    Private metered, unmetered, undefined ' "enums"
    Private hidden, synchronous 'Run params #2 & #3
    Private ForReading, ForWriting 'OpenTextFile param #2
    Private CreateNew 'OpenTextFile param #3
    Private L, L2

    Sub Class_Initialize
        Set sh = CreateObject( "WScript.Shell" )
        Set fso = CreateObject( "Scripting.FileSystemObject" )

        'defaults
        userInteractive = True
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

        With CreateObject( "VBScripting.Includer" )
            Execute .Read( "StringFormatter" )
            Execute .Read( "VBSLogger" )
            Execute .Read( "VBSApp" )
        End With
        Set format = New StringFormatter
        Set log = New VBSLogger
        Set app = New VBSApp
        configFile = format(Array("%AppData%\VBScripting\%s.config", app.GetBaseName))
        ReadConfigFile
   End Sub

    Sub Class_Terminate
        Set sh = Nothing
        Set fso = Nothing
    End Sub
End Class
