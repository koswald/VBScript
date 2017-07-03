
'VBSApp class

'Intended to support identical handling of class procedures by .vbs/.wsf files and .hta files.

'This can be useful when writing a class that might be used in both types of "apps". Note that the VBScript code in the two examples below is identical except for the comments.

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

    'Property GetArgs
    'Returns: array of strings
    'Remark: Returns an array of command-line arguments.
    Property Get GetArgs : GetArgs = arguments : End Property

    'Property GetArgsString
    'Returns: a string
    'Remark: Returns the command-line arguments string. Can be used when restarting a script for example, in order to retain the original arguments. Each argument is wrapped wih quotes, which are stripped off as they are read back in. The return string has a leading space, by design, unless there are no arguments
    Property Get GetArgsString : GetArgsString = argumentsString : End Property

    'Property GetArg
    'Parameter: an integer
    'Returns: a string
    'Remark: Returns the command-line argument having the specified zero-based index.
    Property Get GetArg(index)
        GetArg = ""
        On Error Resume Next
            GetArg = arguments(index)
        On Error Goto 0
    End Property

    'Property GetArgsCount
    'Returns: an integer
    'Remark: Returns the number of arguments.
    Property Get GetArgsCount : GetArgsCount = UBound(arguments) + 1 : End Property

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
    
    'Property GetExe
    'Returns: a string
    'Remark: Returns "mshta.exe" to hta files, and "wscript.exe" or "cscript.exe" to scripts, depending on the host.
    Property Get GetExe
        If IAmAnHta Then
            GetExe = "mshta.exe"
        ElseIf IAmAScript Then
            If "cscript.exe" = LCase(Right(WScript.FullName, 11)) Then GetExe = "cscript.exe"
            If "wscript.exe" = LCase(Right(WScript.FullName, 11)) Then GetExe = "wscript.exe"
        Else
            Err.Raise 3, GetFileName, "Couldn't determine the host .exe; source: VBSApp.vbs::GetExe"
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
    Sub Sleep(milliseconds)
        sh.Run """" & libraryPath & "\VBSApp.wsf"" " & milliseconds, hidden, synchronous
    End Sub
    
    Private oHtaObject_400BFC32009942E895C3F39EA37103DF 'must differ from calling hta's id
    Private fso, sh
    Private filespec, arguments, argumentsString
    Private IAmAnHta, IAmAScript
    Private HtaObjectErrMessage, MissingHtaIdErrMessage
    Private userInteractive, visibility, visible, hidden
    Private synchronous
    Private onUserCancelQuitApp
    Private libraryPath

    Sub Class_Initialize
        HtaObjectErrMessage = "The VBSApp class could not determine the hta file's id. A valid id must be specified in the hta:application element's id property. Source: VBSApp.vbs::"
        MissingHtaIdErrMessage = "An id must be declared in the hta:application element of the .hta file."
        With CreateObject("includer")
            libraryPath = .LibraryPath
        End With
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set sh = CreateObject("WScript.Shell")
        hidden = 0
        visible = 1
        synchronous = True
        SetUserInteractive True
        SetOnUserCancelQuitApp True
        InitializeAppTypes
        arguments = PrivateGetArgs
        argumentsString = PrivateGetArgsString
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
            filespec = Replace(Replace(Replace(document.location.href, "file:///", ""), "%20", " "), "/", "\")
            SetHtaObj GetHtaId
        Else
            Err.Raise 2, GetFileName, "VBSApp.vbs::InitializeFileTypes could not determine the type of file that is calling it."
        End If
    End Sub

    'Private Method SetHtaObj
    'Parameter: the HTA id
    'Remark: Required for .hta files before accessing the arguments properties. The id is defined as a property of the .hta file's hta:application element.
    Private Sub SetHtaObj(id)
        Execute("Set oHtaObject_400BFC32009942E895C3F39EA37103DF = " & id)
    End Sub

    'extract the id from the hta file
    Private Function GetHtaId
        'extract from the file the tag that should contain the id
        With CreateObject("includer")
            Execute(.read("VBSExtracter"))
        End With
        Dim extracter : Set extracter = New VBSExtracter
        extracter.SetFile filespec
        extracter.SetPattern "<hta:application.+id=.+>"
        Dim tag : tag = extracter.extract
        'extract the id from the tag
        Dim re : Set re = New RegExp
        re.Pattern = "id ?= ?""?(\w+)""?" 'quotes are optional, hence the ?
        re.IgnoreCase = True
        Dim matches : Set matches = re.Execute(tag)
        On Error Resume Next
            Dim match : Set match = matches(0)
            GetHtaId = match.Submatches(0)
            If Err Then
                If GetUserInteractive Then If vbOK = MsgBox(MissingHtaIdErrMessage & vbLf & vbLf & filespec & vbLf & vbLf & "Do you want to open the .hta file?", vbExclamation + vbOKCancel, GetFileName) Then sh.Run "notepad """ & GetFullName & """"
                Quit
            End If
        On Error Goto 0
        'release object memory
        Set match = Nothing
        Set matches = Nothing
        Set re = Nothing
    End Function

    'Raise an error if the hta object wasn't properly initialized
    Private Sub EnsureHtaObject(procedure)
        If "HTMLGenericElement" = TypeName(oHtaObject_400BFC32009942E895C3F39EA37103DF) Then
            Exit Sub
        ElseIf "Empty" = TypeName(oHtaObject_400BFC32009942E895C3F39EA37103DF) Then
            Err.Raise 1,, HtaObjectErrMessage & procedure
        End If
    End Sub

    'Private Function GetHtaArgs
    'Returns: an array
    'Remark: Returns the mshta.exe command line args as an array, including the .hta filespec, which has index 0.
    Private Function GetHtaArgs
        EnsureHtaObject("GetHtaArgs")
        Dim cl : cl = oHtaObject_400BFC32009942E895C3F39EA37103DF.CommandLine
        'the command line may contain two spaces between the .hta filespec and the other args. See HKCR\htafile\Shell\Open\Command
        'Todo: make this solution more robust!
        cl = Replace(cl, """ """, """  """)
        Dim args : args = Split(cl, """  """) 'note the double space
        'remove the remaining double quotes
        Dim i : For i = 0 To UBound(args)
            args(i) = Replace(args(i), """", "")
        Next
        GetHtaArgs = args
    End Function

    'Private Property PrivateGetArgs
    'Returns: array of strings
    'Remark: Returns the command-line arguments
    Private Property Get PrivateGetArgs
        With CreateObject("includer")
            Execute(.read("VBSArrays"))
        End With
        Dim arrayUtility : Set arrayUtility = New VBSArrays
        Dim args
        If IAmAnHta Then
            args = GetHtaArgs
            'strip off the first argument, which is the filespec
            PrivateGetArgs = arrayUtility.RemoveFirstElement(args)
        ElseIf IAmAScript Then
            PrivateGetArgs = arrayUtility.CollectionToArray(WScript.Arguments)
        End If
    End Property

    'Private Property PrivateGetArgsString
    'Returns a string
    'Remark: Returns the command-line arguments string with a leading space. For use when restarting a script, in order to retain the original arguments. Each argument is wrapped wih quotes, which are stripped off as they are read back in. The return string has a leading space, by design, unless there are no arguments
    Private Property Get PrivateGetArgsString
        Dim i, s, arg : s = "" : args = GetArgs
        For i = 0 To UBound(args)
            s = s & " """ & args(i) & """"
        Next
        PrivateGetArgsString = s
    End Property

    Private Sub ReleaseObjectMemory
        Set fso = Nothing
        Set sh = Nothing
        Set oHtaObject_400BFC32009942E895C3F39EA37103DF = Nothing
    End Sub

    Sub Class_Terminate
        ReleaseObjectMemory
    End Sub
End Class
