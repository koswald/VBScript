
'VBSApp class

'Intended to support identical handling of class procedures by .vbs/.wsf files and .hta files.

'This can be useful when writing a class that might be used in both types of "apps". Note that the VBScript code in the two examples below is identical except for the comments and indentation.

'' 'test.vbs "arg one" "arg two"
'' With CreateObject("includer")
''     Execute(.read("VBSApp"))
'' End With
'' Dim app : Set app = New VBSApp
'' MsgBox app.GetFullName,, "app.GetFullName" '..\test.vbs
'' MsgBox app.GetArg(1),, "app.GetArg(1)" 'arg two
'' MsgBox app.GetArgsCount,, "app.GetArgsCount" '2
'' app.Quit
'
'' &lt;!--test.hta "arg one" "arg two"-->
'' &lt;hta:application id="oHta" icon="msdt.exe"> &lt;!--an id must be used for command-line args functionality-->
''     &lt;script language="VBScript">
''         With CreateObject("includer")
''             Execute(.read("VBSApp"))
''         End With
''         Dim app : Set app = New VBSApp
''         MsgBox app.GetFullName,, "app.GetFullName" '..\test.hta
''         MsgBox app.GetArg(1),, "app.GetArg(1)" 'arg two
''         MsgBox app.GetArgsCount,, "app.GetArgsCount" '2
''         app.Quit
''     &lt;/script>
'' &lt;/hta:application>
'
Class VBSApp

    'Private Method InitializeHtaDependencies
    'Remark: Initializes members required for .hta files.
    Private Sub InitializeHtaDependencies
        With CreateObject("includer")
            Execute(.read("HTAApp"))
        End With
        Set hta = New HTAApp
    End Sub

    'Property GetArgs
    'Returns: array of strings
    'Remark: Returns an array of command-line arguments.
    Property Get GetArgs
        If Not "Empty" = TypeName(arguments) Then
            GetArgs = arguments
            Exit Property
        End If
        With CreateObject("includer")
            Execute(.read("VBSArrays"))
        End With
        Dim arrayUtility : Set arrayUtility = New VBSArrays
        If IAmAnHta Then
            hta.SetObj hta.GetId
            'strip off the first argument, which is the filespec
            arguments = arrayUtility.RemoveFirstElement(hta.GetArgs)
        ElseIf IAmAScript Then
            arguments = arrayUtility.CollectionToArray(WScript.Arguments)
        End If
        GetArgs = arguments
    End Property

    'Property GetArgsString
    'Returns: a string
    'Remark: Returns the command-line arguments string. Can be used when restarting a script for example, in order to retain the original arguments. Each argument is wrapped wih double quotes. The return string has a leading space, by design, unless there are no arguments.
    Property Get GetArgsString
        If Not "Empty" = TypeName(argumentsString) Then
            GetArgsString = argumentsString
            Exit Property
        End If
        Dim i, s, args : s = "" : args = GetArgs
        For i = 0 To UBound(args)
            s = s & " """ & args(i) & """"
        Next
        argumentsString = s
        GetArgsString = s
    End Property
    
    'Property GetArg
    'Parameter: an integer
    'Returns: a string
    'Remark: Returns the command-line argument having the specified zero-based index.
    Property Get GetArg(index)
        Dim args : args = GetArgs
        GetArg = args(index)
    End Property

    'Property GetArgsCount
    'Returns: an integer
    'Remark: Returns the number of arguments.
    Property Get GetArgsCount
        Dim args : args = GetArgs
        GetArgsCount = UBound(args) + 1
    End Property

    'Property GetFullName
    'Returns: a string
    'Remark: Returns the filespec of the calling script or hta.
    Property Get GetFullName : GetFullName = filespec : End Property

    'Property GetFileName
    'Returns: a string
    'Remark: Returns the name of the calling script or hta, including the filename extension.
    Property Get GetFileName : GetFileName = fso.GetFileName(filespec) : End Property

    'Property GetBaseName
    'Returns: a string
    'Remark: Returns the name of the calling script or hta, without the filename extension.
    Property Get GetBaseName : GetBaseName = fso.GetBaseName(filespec) : End Property

    'Property GetExtensionName
    'Returns: a string
    'Remark: Returns the filename extension of the calling script or hta.
    Property Get GetExtensionName : GetExtensionName = fso.GetExtensionName(filespec) : End Property
    
    'Property GetParentFolderName
    'Returns: a string
    'Remark: Returns the folder that contains the calling script or hta.
    Property Get GetParentFolderName : GetParentFolderName = fso.GetParentFolderName(filespec) : End Property

    'Property GetExe
    'Returns: a string
    'Remark: Returns "mshta.exe" to hta files, and "wscript.exe" or "cscript.exe" to scripts, depending on the host.
    Property Get GetExe
        If IAmAnHta Then
            GetExe = "mshta.exe"
        ElseIf IAmAScript Then
            GetExe = LCase(Right(WScript.FullName, 11))
        Else
            Err.Raise 3, GetFileName, "Couldn't determine the host .exe; source: VBSApp.GetExe"
        End If
    End Property

    'Method RestartIfNotPrivileged
    'Remark: Elevates privileges if they are not already elevated. If userInteractive, first warns user that the User Account Control dialog will open.
    Sub RestartIfNotPrivileged
        With CreateObject("includer")
            Execute(.read("PrivilegeChecker"))
        End With
        Dim pc : Set pc = New PrivilegeChecker
        'if already privileged, skip the rest
        If pc Then Exit Sub
        If userInteractive Then If vbCancel = MsgBox("Restart " & GetFileName & " with elevated privileges?" & vbLf & "(The User Account Control dialog may open.)", vbOKCancel + vbQuestion, GetBaseName) Then Quit
        'start a new instance of this script with elevated privileges
        With CreateObject("Shell.Application")
            .ShellExecute GetExe, """" & GetFullName & """ " & GetArgsString,, "runas"
        End With
        'close the current instance of this script
        Quit
    End Sub

    'Method SetUserInteractive
    'Parameter: boolean
    'Remark: Sets userInteractive value. Setting to True can be useful for debugging. Default is True.
    Sub SetUserInteractive(newUserInteractive)
        userInteractive = newUserInteractive
        If userInteractive Then
            visibility = visible
        Else
            visibility = hidden
        End If
    End Sub

    'Property GetUserInteractive
    'Returns: boolean
    'Remark: Returns the userInteractive setting. This setting also may affect the visibility of selected console windows.
    Property Get GetUserInteractive : GetUserInteractive = userInteractive : End Property

    'Method SetVisiblity
    'Parameter: 0 (hidden) or 1 (normal)
    'Remark: Sets the visibility of selected command windows. SetUserInteractive also affects this setting. Default is True.
    Sub SetVisiblity(newVisibility) : visibility = newVisibility : End Sub

    'Property GetVisibility
    'Returns: 0 (hidden) or 1 (normal)
    'Remark: Returns the current visibility setting. SetUserInteractive also affects this setting.
    Property Get GetVisibility : GetVisibility = visibility : End Property
    
    'Method SetOnUserCancelQuitApp
    'Parameter: boolean
    'Remark: Sets whether the calling app will quit if the user cancels out of a dialog. Default is True.
    Sub SetOnUserCancelQuitApp(newOnUserCancelQuitApp)
        onUserCancelQuitApp = newOnUserCancelQuitApp
    End Sub

    'Property GetOnUserCancelQuitApp
    'Returns: boolean
    'Remark: Retrieves the boolean setting specifying whether the calling app will quit if the user cancels out of a dialog.
    Function GetOnUserCancelQuitApp
        GetOnUserCancelQuitApp = onUserCancelQuitApp
    End Function
    
    'Method Quit
    'Remark: Gracefully closes the hta/script, if allowed by settings.
    Sub Quit
        If Not GetOnUserCancelQuitApp Then Exit Sub
        ReleaseObjectMemory
        If IAmAnHta Then
            Self.close
        ElseIf IAmAScript Then
            WScript.Quit
        End If
    End Sub
    
    'Method Sleep
    'Parameter: an integer
    'Remark: Pauses execution of the script or .hta for the specified number of milliseconds.
    Sub Sleep(ByVal milliseconds)
        If IAmAScript Then
            WScript.Sleep milliseconds
        ElseIf IAmAnHta Then
            hta.Sleep milliseconds
        Else
            Err.Raise 54,, "VBSApp.Sleep: unknown app type."
        End If
    End Sub
   
    Private fso, sh
    Private hta
    Private filespec, arguments, argumentsString
    Private IAmAnHta, IAmAScript
    Private userInteractive, visibility, visible, hidden
    Private synchronous
    Private onUserCancelQuitApp
    Private tmr, EffectiveScriptSleepOverhead, AlwaysPrepareToSleep

    Sub Class_Initialize
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set sh = CreateObject("WScript.Shell")
        hidden = 0
        visible = 1
        synchronous = True
        SetUserInteractive True
        SetOnUserCancelQuitApp True
        InitializeAppTypes
    End Sub
    
    'Determine whether the source file is a script or an hta
    Private Sub InitializeAppTypes
        On Error Resume Next
            Dim x : x = WScript.ScriptName
            If Err Then IAmAnHta = True Else IAmAnHta = False
        On Error Goto 0
        IAmAScript = Not IAmAnHta
        If IAmAScript Then
            filespec = WScript.ScriptFullName
        ElseIf IAmAnHta Then
            InitializeHtaDependencies
            filespec = hta.GetFilespec
        Else
            Err.Raise 2, GetFileName, "VBSApp.InitializeAppTypes could not determine the type of file that is calling it."
        End If
    End Sub

    Private Sub ReleaseObjectMemory
        Set fso = Nothing
        Set sh = Nothing
    End Sub

    Sub Class_Terminate
        ReleaseObjectMemory
    End Sub
End Class
