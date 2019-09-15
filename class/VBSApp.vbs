
'VBSApp class

'Intended to support identical handling of class procedures by .vbs/.wsf files and .hta files.

'This can be useful when writing a class that might be used in both types of "apps".

'Four ways to instantiate

'For .vbs/.wsf scripts,
' <pre>  Dim app : Set app = CreateObject("VBScripting.VBSApp") <br />  app.Init WScript </pre>

'For .hta applications,
' <pre>  Dim app : Set app = CreateObject("VBScripting.VBSApp") <br />  app.Init document </pre>

'If the script may be used in .vbs/.wsf scripts or .hta applications
' <pre>  With CreateObject("VBScripting.Includer") <br />      Execute .read("VBSApp") <br />  End With <br />  Dim app : Set app = New VBSApp </pre>

'Alternate method for both .hta and .vbs/.wsf,
' <pre>  Set app = CreateObject("VBScripting.VBSApp") <br />  If "HTMLDocument" = TypeName(document) Then <br />      app.Init document <br />  Else app.Init WScript <br />  End If </pre>

'Examples
' <pre>  'test.vbs "arg one" "arg two" <br />  With CreateObject("VBScripting.Includer") <br />      Execute .read("VBSApp") <br />  End With <br />  Dim app : Set app = New VBSApp <br />  MsgBox app.GetFileName 'test.vbs <br />  MsgBox app.GetArg(1) 'arg two <br />  MsgBox app.GetArgsCount '2 <br />  app.Quit </pre>
'
' <pre>  &lt;!-- test.hta "arg one" "arg two" --> <br />  &lt;hta:application icon="msdt.exe"> <br />      &lt;script language="VBScript"> <br />          With CreateObject("VBScripting.Includer") <br />              Execute .read("VBSApp") <br />          End With <br />          Dim app : Set app = New VBSApp <br />          MsgBox app.GetFileName 'test.hta <br />          MsgBox app.GetArg(1) 'arg two <br />          MsgBox app.GetArgsCount '2 <br />          app.Quit <br />      &lt;/script> <br />  &lt;/hta:application> </pre>
'
Class VBSApp

    'Private Method InitializeHtaDependencies
    'Remark: Initializes members required for .hta files.
    Private Sub InitializeHtaDependencies
        With CreateObject("VBScripting.Includer")
            Execute .read("HTAApp")
        End With
        Set hta = New HTAApp
    End Sub

    'Property GetArgs
    'Returns: array of strings
    'Remark: Returns an array of command-line arguments.
    Property Get GetArgs
        With CreateObject("VBScripting.Includer")
            Execute .read("VBSArrays")
        End With
        Dim arrayUtility : Set arrayUtility = New VBSArrays
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
    'Remark: Returns the command-line arguments string. Can be used when restarting a script for example, in order to retain the original arguments. Each argument is wrapped wih double quotes. The return string has a leading space, by design, unless there are no arguments.
    Property Get GetArgsString
        Dim i, s, args : s = "" : args = GetArgs
        For i = 0 To UBound(args)
            If wrapAll_ Or InStr(args(i), " ") Then
                s = s & " """ & args(i) & """"
            Else
                s = s & " " & args(i)
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

    'Method RestartWith
    'Parameters: #1: host; #2: switch; #3: elevating
    'Remark: Restarts the script/app with the specified host (typically "wscript.exe", "cscript.exe", or "mshta.exe") and retaining the command-line arguments. Paramater #2 is a cmd.exe switch, "/k" or "/c". Parameter #3 is a boolean, True if restarting with elevated privileges. If userInteractive, first warns user that the User Account Control dialog will open.
    Sub RestartWith(host, switch, elevating)
        Dim format : Set format = CreateObject("VBScripting.StringFormatter")
        Dim privileges : If elevating Then privileges = "runas" Else privileges = ""
        Dim start
        If "cscript" = LCase(fso.GetBaseName(host)) Then
            start = ""
        Else 
            'prevent console window from persisting needlessly
            start = "start"
        End If
        If elevating And userInteractive Then
            Dim msg : msg = format(Array( _
                " Restart %s with elevated privileges? %s (The User Account Control dialog may open.)", _
                GetFileName, vbLf _
            ))
            If vbCancel = MsgBox(msg, vbOKCancel + vbQuestion, GetBaseName) Then
                Quit
            End If
        End If
        Dim args : args = format(Array( _
            "%s cd ""%s"" & %s %s ""%s"" %s", _
             switch, GetParentFolderName, start, host, me.GetFullName, GetArgsString _
        ))
        With CreateObject("Shell.Application")
            .ShellExecute "cmd", args, GetParentFolderName, privileges
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

    'Method SetVisibility
    'Parameter: 0 (hidden) or 1 (normal)
    'Remark: Sets the visibility of selected command windows. SetUserInteractive also affects this setting. Default is True.
    Sub SetVisibility(newVisibility) : visibility = newVisibility : End Sub

    'Property GetVisibility
    'Returns: 0 (hidden) or 1 (normal)
    'Remark: Returns the current visibility setting. SetUserInteractive also affects this setting.
    Property Get GetVisibility : GetVisibility = visibility : End Property
    
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
            Err.Raise 54,, "VBSApp.Sleep: unknown app type."
        End If
    End Sub
    
    'Property WScriptHost
    'Returns: "wscript.exe"
    'Remark: Can be used as an argument for the method RestartIfNotPrivileged.
    Public Property Get WScriptHost : WScriptHost = "wscript.exe" : End Property
    
    'Property CScriptHost
    'Returns: "cscript.exe"
    'Remark: Can be used as an argument for the method RestartIfNotPrivileged.
    Public Property Get CScriptHost : CScriptHost = "cscript.exe" : End Property
    
    'Property GetHost
    'Returns: "wscript.exe" or "cscript.exe" or "mshta.exe"
    'Remark: Returns the current host. Can be used as an argument for the method RestartIfNotPrivileged.
    Public Property Get GetHost : GetHost = GetExe : End Property
    
    'Determine whether the source file is a script,
    '(.wsf or .vbs), an .hta, or a .wsc
    Sub InitializeAppTypes
        IAmAScript = False : IAmAnHta = False : IAmAWsc = False
        If "HTMLDocument" = TypeName(document) Then
            IAmAnHta = True
            InitializeHtaDependencies
            filespec = hta.GetFilespec
        ElseIf "Object" = TypeName(WScript) Then
            IAmAScript = True
            filespec = WScript.ScriptFullName
        ElseIf "VBSApp" = TypeName(Me) Then
            IAmAWsc = True
        Else
            Err.Raise 2, GetFileName, "VBSApp.InitializeAppTypes could not determine the type of file that is calling it."
        End If
    End Sub
    
    Property Get WrapAll : WrapAll = wrapAll_ : End Property
    Property Let WrapAll(newWrapAll)
        wrapAll_ = newWrapAll
    End Property

    Private fso
    Private hta
    Private wrapAll_
    Private filespec, arguments, argumentsString
    Private IAmAnHta, IAmAScript, IAmAWsc
    Private userInteractive, visibility, visible, hidden

    Sub Class_Initialize
        Set fso = CreateObject("Scripting.FileSystemObject")
        hidden = 0
        visible = 1
        wrapAll = False
        SetUserInteractive True
        InitializeAppTypes
    End Sub

    Private Sub ReleaseObjectMemory
        Set fso = Nothing
    End Sub

End Class
