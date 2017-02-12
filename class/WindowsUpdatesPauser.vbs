
'Pause Windows Updates to get more bandwidth. Don't forget to resume.

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

    Private sh, fso, format 'objects
    Private configFile, thisFile
    Public appName
    Private profileName, srvcType 'data kept in .config file
    Private metered, unmetered 'enumish
    Private hidden, synchronous 'Run params #2 & #3
    Private ForReading 'OpenTextFile param #2

    Sub Class_Initialize
        Set sh = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")
        unmetered = "Unmetered" ' "enums"
        metered = "Metered"
        hidden = 0 'constants
        synchronous = True
        ForReading = 1
        Dim reader : Set reader = CreateObject("includer")
        Execute(reader.read("StringFormatter"))
        Set format = New StringFormatter
        Set reader = Nothing
        On Error Resume Next
            thisFile = Replace(Replace(document.location.href, "file:///", ""), "%20", " ") 'called by .hta
            If Err Then thisFile = WScript.ScriptFullName 'called by script
        On Error Goto 0
        appName = fso.GetBaseName(thisFile)
        configFile = appName & ".config"
        ReadConfigFiles
   End Sub

    'Sub ReadConfigFiles
    'Remark: Read profileName, srvcType, and configFile from .config file(s)

    Private Sub ReadConfigFiles
        sh.CurrentDirectory = fso.GetParentFolderName(thisFile) 'make the current directory the parent folder of this file, so that the first place to look for a .config file doesn't depend on a shortcut's working directory
        ExecuteConfigFile
        ExecuteConfigFile 'execute again in case the first configFile points to another configFile
    End Sub

    Private Sub ExecuteConfigFile
        On Error Resume Next
            Dim stream : Set stream = fso.OpenTextFile(sh.ExpandEnvironmentStrings(configFile), ForReading)
            If Err Then MsgBox "Couldn't find " & configFile, vbCritical, appName
        On Error Goto 0
        Execute(stream.ReadAll)
        stream.Close
        Set stream = Nothing
    End Sub

    Sub ShowStatusError
        Dim L, L2 : L = vbLf : L2 = L & L
        MsgBox format(Array( _
            "The metered status of the connection %s could not be determined. %s" & _
            "Check the profileName (network name) in %s" & _
            "%s will now open the .config file." _
            , L2 & profileName & L2, L, L2 & configFile & L2, appName)) _
            , vbCritical, appName
        sh.Run "notepad """ & configFile & """"
    End Sub

    Sub Class_Terminate
        Set sh = Nothing
        Set fso = Nothing
    End Sub
End Class
