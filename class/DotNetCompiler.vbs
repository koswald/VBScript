
'A wrapper for .NET compiler functions

'Visual Studio is recommended but not required.
'Includes generating a strong name key pair, compiling a C# or VB file with csc.exe, and registering the .dll.

'Note: Normally you would want two .dll files, one compiled and registered with 32-bit .exe files (csc.exe and RegAsm.exe), and the other compiled and registered with 64-bit .exe files. The two .dll files would be named differently and/or kept in different directories. If however, you want to to just use one .dll, and you want it to be available to 64-bit and 32-bit processes, then it must be compiled using the 32-bit csc.exe, and you must register it twice, once with the 32-bit RegAsm.exe and again with the 64-bit RegAsm.exe, changing from one .exe to the other using the SetBitness method.
'
'<a href="https://msdn.microsoft.com/en-us/library/78f4aasd.aspx">Command-line Building With csc.exe</a>
'
Class DotNetCompiler

    'Method RestartIfNotPrivileged
    'Parameters: #1: "wscript.exe", "cscript.exe", or "mshta.exe"; #2: "/k" or "/c" (cmd.exe switch)
    'Remark: Elevates privileges if they are not already elevated. If app.GetUserInteractive is True, the user is first warned that the User Account Control dialog will open.
    Sub RestartIfNotPrivileged(host, switch)
        app.RestartWith host, switch, False
    End Sub

    'Method SetTargetFolder
    'Parameter: targetFolder
    'Remark: Sets the destination folder for the .dll, .exe. Optional. Default is the folder containing the calling script.
    Sub SetTargetFolder(newTargetFolder)
        targetFolder = fso.GetAbsolutePathName(newTargetFolder)
        vfs.MakeFolder(targetFolder)
    End Sub

    'Method SetTargetName
    'Parameter: targetName
    'Remark: Sets the base name of the .snk, .dll, or .exe file that will be created. E.g. To create a .dll named Vox32.dll, use "Vox32". E.g. To generate a .snk key pair named Vox.snk, use "Vox". Optional, if a filespec or file name is passed in on the command line. Required otherwise.
    Sub SetTargetName(newTargetName) : targetName = newTargetName : End Sub

    'Property GetTargetName
    'Returns targetName
    Property Get GetTargetName : GetTargetName = targetName : End Property

    'Method SetSourceFile
    'Parameter: sourceFile
    'Remark: Sets the filespec of the source file. May use a relative path, relative to the location of the calling script. Optional if a valid file is specified on the command line. The .snk file will be created in the same folder.
    Sub SetSourceFile(newSourceFile)
        If "" = newSourceFile Then sourceFile = "" : Exit Sub
        sourceFile = fso.GetAbsolutePathName(newSourceFile)
        If Not fso.FileExists(sourceFile) Then
            'couldn't find the source file; this can happen when
            'elevating privileges changes the current directory
            'to C:\Windows\System32; so assume it's a relative path,
            'and try using the script's location as the reference
            'for the relative path
            sh.CurrentDirectory = fso.GetParentFolderName(app.GetFullName)
            sourceFile = fso.GetAbsolutePathName(newSourceFile)
        End If
        SetTargetName fso.GetBaseName(sourceFile)
        sourceFolder = fso.GetParentFolderName(sourceFile)
    End Sub

    'Method SetTargetType
    'Parameter: newTargetType
    'Remark: Sets the target type for the compiled file. Should be library, exe (console app), winexe, or module. Optional. Default is library.
    Sub SetTargetType(newTargetType)
        targetType = newTargetType
        If "library" = LCase(targetType) Then
            ext = "dll"
        ElseIf "module" = LCase(targetType) Then
            ext = "module"
        ElseIf "winexe" = LCase(targetType) Or "exe" = LCase(targetType) Then
            ext = "exe"
        Else
            Err.Raise 93,, app.GetFileName & ": Unexpected target type " & targetType
        End If
    End Sub

    'Method SetUserInteractive
    'Parameter: boolean
    'Remark: Sets userInteractive value. Setting to True can be useful for debugging. Default is False. Using the command-line argument -debug sets this value to True.
    Sub SetUserInteractive(newUserInteractive)
        app.SetUserInteractive newUserInteractive
    End Sub

    'Property GetUserInteractive
    'Returns: a boolean
    'Remark: Returns the userInteractive value.
    Property Get GetUserInteractive : GetUserInteractive = app.GetUserInteractive : End Property

    'Method GenerateKeyPair
    'Remark: Generates a strong name key pair using Visual Studio's sn command. Requires Visual Studio to be installed. Requires the SetKeyFile(keyFile) method to have been called to set the location and filename of the key file.
    Sub GenerateKeyPair

        'validate
        ValidateVisualStudio
        Dim msg
        If "Empty" = TypeName(keyFile) Then
            msg = "Use the SetKeyFile(keyFile) method to designate the key file location and name prior to calling GenerateKeyFile. Environment variables are allowed."
            If "wscript.exe" = LCase(Right(WScript.FullName, 11)) Then sh.PopUp msg, 60, app.GetBaseName, vbSystemModal + vbExclamation
            If "cscript.exe" = LCase(Right(WScript.FullName, 11)) Then Err.Raise 254,, msg
            Exit Sub
        End If
        Dim parent : parent = Expand(fso.GetParentFolderName(fso.GetAbsolutePathName(keyFile)))
        vfs.MakeFolder(parent)

        'build the arguments
        args = format(Array( _
            "/c @echo on & ""%s"" & sn -k ""%s""", _
            batFile, keyFile _
        ))
        If app.GetUserInteractive Then args = args & " & echo. & pause"

        'give an opt out
        msg = "Verify arguments for key generation"
        If app.GetUserInteractive Then If vbCancel = MsgBox(args, vbOKCancel, msg & " - " & scriptName) Then Quit

        'generate the strong name key pair
        sh.Run "%ComSpec% " & args, app.GetVisibility, synchronous

    End Sub

    'initialize the variable batFile
    'using values from DotNetCompilier.config
    Private Sub ValidateVisualStudio
        Dim i
        For i = 1 To UBound(batFiles)
            batFile = Expand(batFiles(i))
            If fso.FileExists(batFile) Then Exit Sub
        Next
        Err.Raise 25, scriptName, "Couldn't find any of the batch files configured in DotNetCompiler.config, which enable the Visual Studio strong name tool. If you don't have Visual Studio, but you have some version of the .NET framework in C:\Windows\Microsoft.NET\Framework, then you can still compile and register without a strong name: Remove or comment out the line in the .cs file with AssemblyKeyFileAttribute. You will receive a warning message when registering the .dll."
    End Sub

    'expand environment variables
    Function Expand(str)
        Expand = sh.ExpandEnvironmentStrings(str)
    End Function

    Private Sub ValidateName
        If "" = targetName Then SetTargetName fso.GetBaseName(sourceFile)
        If "" = targetName Then Err.Raise 63, scriptName, scriptName & " needs a name. Use SetTargetName <targetName>, or pass in the name or the source file on the command line."
    End Sub

    'initialize the variables exeFolder32 and exeFolder64
    'using values from DotNetCompiler.config
    Private Sub ValidateDotNetFolders
        Dim i, folders
        For i = 1 To UBound(exeFolders)
            folders = Split(exeFolders(i), "|")
            exeFolder64 = Trim(Expand(folders(0)))
            exeFolder32 = Trim(Expand(folders(1)))
            If fso.FolderExists(exeFolder64) And fso.FolderExists(exeFolder32) Then Exit Sub
        Next
        Err.Raise 1, scriptName, "Couldn't verify the location of any of the .NET executables folder pairs in DotNetCompiler.config."
    End Sub

    'Method Compile
    'Remark: Uses csc.exe to compile the .cs or .vb file specified by SetSourceFile or passed in on the command line.
    Sub Compile

        'validate
        If Not fso.FileExists(sourceFile) Then Err.Raise 1, scriptName, "Couldn't find the source file """ & sourceFile & """. Use SetSourceFile sourceFile, or pass in the source file on the command line."
        If Not "cs" = LCase(fso.GetExtensionName(sourceFile)) Then Err.Raise 1, scriptName, "A .cs file is required for compiling."
        ValidateName

        'build the command arguments
        Dim args : args = args & " /target:" & targetType
        If supressWarnings Then args = args & " /warn:0"
        If debug Then args = args & " /debug"
        If Not "Empty" = TypeName(keyFile) Then
            args = args & " /keyfile:""" & keyFile & """"
        Else
            args = args & " /delaysign"
        End If

        'build the commmand string
        Dim cmd : cmd = format(Array( _
            "%ComSpec% /c cd ""%s"" & echo. & ""%s\csc.exe"" /out:%s.%s %s %s ""%s"" & echo.", _
            sourceFolder, exeFolder, targetName, ext, refs, args, sourceFile _
        ))
        If app.GetUserInteractive Then cmd = cmd & " & pause"

        'give an opt out
        msg = "Verify command for compiling"
        If app.GetUserInteractive Then If vbCancel = MsgBox(cmd, vbOKCancel, msg & " - " & scriptName) Then Quit

        'compile
        sh.Run cmd, app.GetVisibility, synchronous
    End Sub

    'Method SetSupressWarnings
    'Parameter: boolean
    'Remark: Sets the csc.exe compiler to supress warnings, if True is specified. Optional. Default is False.
    Sub SetSupressWarnings(newSupressWarnings) : supressWarnings = newSupressWarnings : End Sub

    'Method SetDebug
    'Parameter: boolean
    'Remark: Sets the csc.exe compiler to debug mode, if True is specified. Optional. Default is False.
    Sub SetDebug(newDebug) : debug = newDebug : End Sub

    'Method AddRef
    'Parameter: ref
    'Remark: Adds the specified assembly reference, a filespec, to the csc.exe compiler command prior to calling the Compile method. Optional.
    Sub AddRef(ref)
        refs = refs & " /r:""" & ref & """"
    End Sub

    'Method Unregister
    'Remark: Uses RegAsm.exe to unregister the .dll file.
    Sub Unregister
        unregistering = True
        Register
    End Sub

    'Method Register
    'Remark: Uses RegAsm.exe to register the .dll file.
    Sub Register

        'validate
        If Not pc.Privileged Then Err.Raise 1, scriptName, "Registering a .dll requires elevated privileges."
        ValidateName

        'move the .dll to the target folder
        Dim sourceDllFile : sourceDllFile = sourceFolder & "\" & targetName & ".dll"
        Dim targetDllFile : targetDllFile = targetFolder & "\" & targetName & ".dll"
        If Not vfs.FoldersAreTheSame(sourceFolder, targetFolder) And Not unregistering Then
            If fso.FileExists(targetDllFile) Then
                vfs.DeleteFile targetDllFile
            End If
            fso.MoveFile sourceDllFile, targetDllFile
        End If

        'build the argument(s)
        Dim args : args = format(Array( _
            "/c cd ""%s"" & echo. & ""%s\RegAsm.exe"" %s.dll", _
            targetFolder, exeFolder, targetName _
        ))
        'args = args & " /tlb:" & targetName & ".tlb" 'create a type library
        args = args & " /codebase" 'if not putting .dll in the GAC
        If unregistering Then args = args & " /unregister"
        If app.GetUserInteractive Then args = args & " & echo. & pause "

        'give an opt out
        Dim action : If unregistering Then action = "unregistering" Else action = "registering"
        msg = "Verify arguments for " & action
        If app.GetUserInteractive Then If vbCancel = MsgBox(args, vbOKCancel, msg & " - " & scriptName) Then Quit

        'register or unregister
        sh.Run "%ComSpec% " & args, app.GetVisibility, synchronous
    End Sub

    'Method SetBitness
    'Parameter: 64 or 32
    'Remark: Sets the bitness for the compiler or registrar to 64 or 32. Optional. Default is 64.
    Sub SetBitness(bitness)
        ValidateDotNetFolders
        If 64 = bitness Then
            exeFolder = exeFolder64
        ElseIf 32 = bitness Then
            exeFolder = exeFolder32
        End If
    End Sub

    'Method SetKeyFile
    'Parameter: keyFile
    'Remark: Sets the location and name of the stong-name key pair file (.snk).
    Sub SetKeyFile(newKeyFile)
        keyFile = fso.GetAbsolutePathName(newKeyFile)
    End Sub
        
    'Method Quit
    'Remark: Gracefully quits the hta/script, if allowed by settings.
    Sub Quit
        ReleaseObjectMemory
        app.Quit
    End Sub

    Private fso 'Scripting.FileSystemObject
    Private sh 'WScript.Shell object
    Private sa 'Shell.Application object
    Private vfs 'VBSFileSystem object
    Private pc 'PrivilegeChecker object
    Private app 'VBSApp object
    Private format 'StringFormatter object
    Private batFile, batfiles
    Private exeFolders, exeFolder, exeFolder64, exeFolder32
    Private sourceFile, sourceFolder, targetFolder, targetName
    Private ext, targetType
    Private scriptName, thisFile
    Private refs, args, cmd
    Private supressWarnings, debug 'compiler settings
    Private unregistering 'registration setting
    Private msg, L
    Private synchronous, visible, hidden
    Private keyFile

    Sub Class_Initialize
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set sh = CreateObject("WScript.Shell")
        Set sa = CreateObject("Shell.Application")
        With CreateObject("includer")
            Execute .read("VBSFileSystem")
            Execute .read("PrivilegeChecker")
            Execute .read("VBSApp")
            Execute(.read("DotNetCompiler.config"))
            Execute .read("StringFormatter")
        End With
        Set vfs = New VBSFileSystem
        Set pc = New PrivilegeChecker
        Set app = New VBSApp
        Set format = New StringFormatter
        'get the filespec of the calling script
        thisFile = app.GetFullName
        SetSourceFile ""
        On Error Resume Next
            SetSourceFile app.GetArg(0)
        On Error Goto 0
        scriptName = fso.GetFileName(thisFile)
        L = vbLf & vbTab
        synchronous = True
        hidden = 0
        visible = 1
        unregistering = False 'defaults
        refs = ""
        SetTargetType "library"
        SetSupressWarnings False
        SetDebug False
        SetBitness 64
        sh.CurrentDirectory = fso.GetParentFolderName(thisFile)
        SetTargetFolder sh.CurrentDirectory
        SetTargetName fso.GetBaseName(sourceFile)
        app.SetUserInteractive False
        'add references from the command-line, if any
        Dim args, nextArg : args = app.GetArgs
        If UBound(args) > 0 Then
            For i = 0 To UBound(args)
                nextArg = ""
                On Error Resume Next
                    nextArg = app.GetArg(i + 1)
                On Error Goto 0
                If "-ref" = LCase(app.GetArg(i)) Then If fso.FileExists(nextArg) Then AddRef nextArg
                If "-debug" = LCase(app.GetArg(i)) Then app.SetUserInteractive True
            Next
        End If

    End Sub

    Private Sub ReleaseObjectMemory
        Set fso = Nothing
        Set sa = Nothing
        Set sh = Nothing
    End Sub

    Sub Class_Terminate
        ReleaseObjectMemory
    End Sub
End Class
