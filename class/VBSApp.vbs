
'VBSApp class

'Intended to support identical handling of class procedures by .vbs/.wsf files and .hta files.

'This can be useful when writing a class that might be used in both types of "apps".

'Four ways to instantiate

'For .vbs/.wsf scripts,
' <pre>  Dim app : Set app = CreateObject( "VBScripting.VBSApp" ) <br />  app.Init WScript </pre>

'For .hta applications,
' <pre>  Dim app : Set app = CreateObject( "VBScripting.VBSApp" ) <br />  app.Init document </pre>

'If the script may be used in .vbs/.wsf scripts or .hta applications
' <pre>  With CreateObject( "VBScripting.Includer" ) <br />      Execute .Read( "VBSApp" ) <br />  End With <br />  Dim app : Set app = New VBSApp </pre>

'Alternate method for both .hta and .vbs/.wsf,
' <pre>  Set app = CreateObject( "VBScripting.VBSApp" ) <br />  If "HTMLDocument" = TypeName(document) Then <br />      app.Init document <br />  Else app.Init WScript <br />  End If </pre>

'Examples
' <pre>  'test.vbs "arg one" "arg two" <br />  With CreateObject( "VBScripting.Includer" ) <br />      Execute .Read( "VBSApp" ) <br />  End With <br />  Dim app : Set app = New VBSApp <br />  MsgBox app.GetFileName 'test.vbs <br />  MsgBox app.GetArg(1) 'arg two <br />  MsgBox app.GetArgsCount '2 <br />  app.Quit </pre>
'
' <pre>  &lt;!-- test.hta "arg one" "arg two" --> <br />  &lt;hta:application icon="msdt.exe"> <br />      &lt;script language="VBScript"> <br />          With CreateObject( "VBScripting.Includer" ) <br />              Execute .Read( "VBSApp" ) <br />          End With <br />          Dim app : Set app = New VBSApp <br />          MsgBox app.GetFileName 'test.hta <br />          MsgBox app.GetArg(1) 'arg two <br />          MsgBox app.GetArgsCount '2 <br />          app.Quit <br />      &lt;/script> <br />  &lt;/hta:application> </pre>
'
Class VBSApp

    Private fso 'Scripting.FileSystemObject
    Private hta 'VBSHta object
    Private incl 'VBScripting.Includer object
    Private wrapAll_ 'holds WrapAll property value
    Private filespec, argumentsString 'strings
    Private arguments 'array of strings
    Private IAmAnHta, IAmAScript 'booleans
    Private userInteractive 'boolean
    Private visibility, visible, hidden 'integers
    Private powershell 'pwsh.exe filespec or just "powershell"
    Private UACMsg 'User Account Control warning message
    
    Sub Class_Initialize
        Set fso = CreateObject( "Scripting.FileSystemObject" )
        Set incl = CreateObject( "VBScripting.Includer" )
        hidden = 0
        visible = 1
        wrapAll = False
        SetUserInteractive True
        IAmAScript = False
        IAmAnHta = False
        InitializeAppTypes
        Execute incl.Read( "Configurer" )
        With New Configurer
            powershell = .PowerShell
        End With
        UACMsg = " Restart %s with elevated privileges? %s (The User Account Control dialog may open.)"
    End Sub

    'Determine whether the source file is a script,
    '(.wsf or .vbs), an .hta, or a .wsc. This method is for when the VBSApp object is included by 1) direct reference in a <script> tag in a an .hta or .wsf file or 2) by an Execute includer.Read() statement.
    Sub InitializeAppTypes
        If "HTMLDocument" = TypeName(document) Then
            IAmAnHta = True
            InitializeHtaDependencies
            filespec = hta.GetFilespec
        ElseIf "Object" = TypeName(WScript) Then
            IAmAScript = True
            filespec = WScript.ScriptFullName
        End If
    End Sub

    'Private Method InitializeHtaDependencies
    'Remark: Initializes members required for .hta files.
    Private Sub InitializeHtaDependencies
        Execute incl.Read( "HTAApp" )
        Set hta = New HTAApp
    End Sub

    'Property GetArgs
    'Returns: array of strings
    'Remark: Returns an array of command-line arguments.
    Property Get GetArgs
        Dim arrayUtility
        Execute incl.Read( "VBSArrays" )
        Set arrayUtility = New VBSArrays
        If IAmAnHta Then
            'strip off the first argument, which is the filespec
            arguments = arrayUtility.RemoveFirstElement(hta.GetArgs)
        ElseIf IAmAScript Then
            arguments = arrayUtility.CollectionToArray(WScript.Arguments)
        End If
        GetArgs = arguments
    End Property

    'Property GetArgsString
    'Returns: a string
    'Remark: Returns the command-line arguments string. Can be used when restarting a script for example, in order to retain the original arguments. Arguments are wrapped with double quotes, if they contain spaces or if WrapAll is set to True. The return string has a leading space, by design, unless there are no arguments.
    Property Get GetArgsString
        Dim i, s, args : s = "" : args = GetArgs
        For i = 0 To UBound(args)
            If WrapAll Or InStr( args(i), " " ) Then
                s = s & " """ & args(i) & """"
            Else s = s & " " & args(i)
            End If
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
    Property Get GetFullName
        GetFullName = filespec
    End Property

    'Property GetFileName
    'Returns: a string
    'Remark: Returns the name of the calling script or hta, including the filename extension.
    Property Get GetFileName
        GetFileName = fso.GetFileName(filespec)
    End Property

    'Property GetBaseName
    'Returns: a string
    'Remark: Returns the name of the calling script or hta, without the filename extension.
    Property Get GetBaseName
        GetBaseName = fso.GetBaseName(filespec)
    End Property

    'Property GetExtensionName
    'Returns: a string
    'Remark: Returns the filename extension of the calling script or hta.
    Property Get GetExtensionName
        GetExtensionName = fso.GetExtensionName(filespec)
    End Property

    'Property GetParentFolderName
    'Returns: a string
    'Remark: Returns the folder that contains the calling script or hta.
    Property Get GetParentFolderName
        GetParentFolderName = fso.GetParentFolderName(filespec)
    End Property

    'Property GetExe
    'Returns: a string
    'Remark: Returns "mshta.exe" to hta files, and "wscript.exe" or "cscript.exe" to scripts, depending on the host.
    Property Get GetExe
        If IAmAnHta Then
            GetExe = "mshta.exe"
        ElseIf IAmAScript Then
            GetExe = LCase(Right(WScript.FullName, 11))
        Else
            Err.Raise 17, GetFileName, "Couldn't determine the host .exe; source: VBSApp.GetExe"
        End If
    End Property

    'Method RestartWith
    'Parameters: #1: host; #2: switch; #3: elevating
    'Remark: <strong> Deprecated</strong> in favor of the RestartUsing method. Restarts the script/app with the specified host (typically "wscript.exe", "cscript.exe", or "mshta.exe"), retaining the command-line arguments. Uses cmd.exe for the shell. Parameter #2 is a cmd.exe switch, "/k" or "/c". Parameter #3 is a boolean, True if restarting with elevated privileges. If userInteractive, first warns user that the User Account Control dialog will open.
    Sub RestartWith( host, switch, elevating )
        Dim format 'VBScripting.StringFormatter obj
        Dim start 'string
        Dim msg, settings, title 'MsgBox arguments
        Dim cmd 'string: ShellExecute arg #1
        Dim args 'string: ShellExecute arg #2
        Dim pwd 'string: ShellExecute arg #3
        Dim privileges 'string: ShellExecute arg #4
        Dim hostBaseName 'string: partial filespec: e.g. cscript

        Set format = CreateObject( "VBScripting.StringFormatter" )
        hostBaseName = LCase(fso.GetBaseName(host))

        'Opt out

        If elevating And userInteractive Then
            msg = format(Array( _
                UACMsg, GetFileName, vbLf _
            ))
            settings = vbOKCancel + vbQuestion
            title = GetFileName
            If vbCancel = MsgBox( msg, settings, title ) Then
                Quit
            End If
        End If

        'Restart the script/hta

        cmd = "cmd"
        If "cscript" = hostBaseName Then
            start = ""
        Else 'prevent console window from persisting
            start = "start"
        End If
        args = format(Array( _
            "%s cd ""%s"" & %s %s ""%s"" %s", _
             switch, GetParentFolderName, start, _
             host, me.GetFullName, GetArgsString _
        ))
        pwd = GetParentFolderName
        If elevating Then
            privileges = "runas"
        Else privileges = ""
        End If
        With CreateObject( "Shell.Application" )
            .ShellExecute cmd, args, pwd, privileges
        End With

        'close the current instance of the script/hta
        Quit
    End Sub

    'Method RestartUsing
    'Parameters: #1: host; #2: exit?; #3: elevate?
    'Remark: Restarts the script/hta with the specified host, "wscript.exe", "cscript.exe", "mshta.exe", or a full path to one of these, retaining the command-line arguments. Uses pwsh.exe for the shell, if available, or falls back to powershell.exe. Unusual or custom paths for pwsh.exe can be specified in the file <code>.configure</code> in the project root folder. Parameter #2 is a boolean specifying whether the powershell window should exit after completion. Parameter #3 is a boolean, True if restarting with elevated privileges. If userInteractive, first warns user that the User Account Control dialog will open. If it is desired to elevate privileges, and privileges are already elevated, and the desired host is already hosting, then the script does not restart: The calling script or hta does not have to check whether privileges are elevated or explicitly call the Quit method.
    Sub RestartUsing( host, exiting, elevating )
        Dim pc 'PrivilegeChecker object
        Dim format 'VBScripting.StringFormatter object
        Dim msg, settings, title 'MsgBox arguments
        Dim params 'powershell parameters
        Dim cmd 'string: ShellExecute arg #1
        'Class scope: args_ 'string: ShellExecute arg #2
        Dim pwd 'string: ShellExecute arg #3
        Dim privileges 'string: ShellExecute arg #4
        Dim hostFileName 'string: partial filespec: e.g. cscript.exe

        Set format = CreateObject( "VBScripting.StringFormatter" )
        hostFileName = LCase(fso.GetFileName(host))
        Execute incl.Read( "PrivilegeChecker" )
        Set pc = New PrivilegeChecker
        If elevating And pc _
        And GetExe = hostFileName _
        And Not RUArgsTest Then
            'privileges are already elevated,
            'desired host is already hosting
            Exit Sub
        ElseIf Not elevating _
        And GetExe = hostFileName _
        And Not RUArgsTest Then
            Exit Sub
        End If

        'Opt out

        If elevating And userInteractive Then
            msg = format( Array( _
                UACMsg, GetFileName, vbLf _
            ))
            settings = vbOKCancel + vbQuestion
            title = GetFileName
            If vbCancel = MsgBox(msg, settings, title) Then
                Quit
            End If
        End If

        'Restart the script/hta

        cmd = powershell
        params = "-ExecutionPolicy Bypass"
        If Not exiting Then
            params = params & " -NoExit"
        End If
        params = params & " -Command"
        args_ = format(Array( _
            "%s Set-Location '%s' ; %s ""'%s'"" %s", _
             params, GetParentFolderName, _
             Expand( host ), me.GetFullName, GetArgsString _
        ))
        pwd = GetParentFolderName
        If elevating Then
            privileges = "runas"
        Else privileges = ""
        End If
        If RUArgsTest Then
            Exit Sub
        End If
        With CreateObject( "Shell.Application" )
            .ShellExecute cmd, args_, pwd, privileges
        End With

        'close the current instance of the script
        Quit
    End Sub

    'For testability for the RestartUsing (RU) method: 
    Private args_
    Property Get RUArgs
        RUArgs = args_
    End Property
    Property Let RUArgsTest( newBoolean )
        ruArgsTest_ = newBoolean
    End Property
    Property Get RUArgsTest
        If IsEmpty( ruArgsTest_ ) Then
            ruArgsTest_ = False
        End If
        RUArgsTest = ruArgsTest_
    End Property
    Private ruArgsTest_

    'Property DoExit
    'Returns True
    'Remark: Suitable for use with the RestartUsing method, argument #2
    Property Get DoExit
        DoExit = True
    End Property

    'Property DoNotExit
    'Returns False
    'Remark: Suitable for use with the RestartUsing method, argument #2
    Property Get DoNotExit
        DoNotExit = False
    End Property

    'Property DoElevate
    'Returns True
    'Remark: Suitable for use with the RestartUsing method, argument #3
    Property Get DoElevate
        DoElevate = True
    End Property

    'Property DoNotElevate
    'Returns False
    'Remark: Suitable for use with the RestartUsing method, argument #3
    Property Get DoNotElevate
        DoNotElevate = False
    End Property

    Function Expand( s )
        With CreateObject( "WScript.Shell" )
            Expand = .ExpandEnvironmentStrings( s )
        End With
    End Function

    'Method SetUserInteractive
    'Parameter: boolean
    'Remark: Sets userInteractive value. Setting to True can be useful for debugging. Default is True.
    Sub SetUserInteractive(newUserInteractive)
        userInteractive = newUserInteractive
        If userInteractive Then
            visibility = visible
        Else visibility = hidden
        End If
    End Sub

    'Property GetUserInteractive
    'Returns: boolean
    'Remark: Returns the userInteractive setting. This setting also may affect the visibility of selected console windows.
    Property Get GetUserInteractive
        GetUserInteractive = userInteractive
    End Property

    'Method SetVisibility
    'Parameter: 0 (hidden) or 1 (normal)
    'Remark: Sets the visibility of selected command windows. SetUserInteractive also affects this setting. Default is 1.
    Sub SetVisibility(newVisibility)
        visibility = newVisibility
    End Sub

    'Property GetVisibility
    'Returns: 0 (hidden) or 1 (normal)
    'Remark: Returns the current visibility setting. SetUserInteractive also affects this setting.
    Property Get GetVisibility
        GetVisibility = visibility
    End Property

    'Method Quit
    'Remark: Gracefully closes the hta/script.
    Sub Quit
        ReleaseObjectMemory
        If IAmAnHta Then
            document.parentWindow.close
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
            Err.Raise 17,, "VBSApp.Sleep: unknown app type."
        End If
    End Sub

    'Property WScriptHost
    'Returns: "wscript.exe"
    'Remark: Can be used as an argument for the method RestartWith.
    Public Property Get WScriptHost
        WScriptHost = "wscript.exe"
    End Property

    'Property CScriptHost
    'Returns: "cscript.exe"
    'Remark: Can be used as an argument for the method RestartWith.
    Public Property Get CScriptHost
        CScriptHost = "cscript.exe"
    End Property

    'Property GetHost
    'Returns: "wscript.exe" or "cscript.exe" or "mshta.exe"
    'Remark: Returns the current host. Can be used as an argument for the method RestartWith.
    Public Property Get GetHost
        GetHost = GetExe
    End Property

    'Boolean: whether to wrap all arguments in quotes, as opposed to just those arguments that contain spaces. See property GetArgsString
    Property Get WrapAll
        WrapAll = wrapAll_
    End Property
    Property Let WrapAll(newWrapAll)
        wrapAll_ = newWrapAll
    End Property

    Private Sub ReleaseObjectMemory
        Set fso = Nothing
    End Sub

End Class
