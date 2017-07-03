
'A wrapper for .NET compiler functions

'Visual Studio is recommended but not required.
'Includes generating a strong name key pair, compiling a C# or VB file with csc.exe, and registering the .dll.

'Note: Normally you would want two .dll files, one compiled and registered with 32-bit .exe files (csc.exe and RegAsm.exe), and the other compiled and registered with 64-bit .exe files. The two .dll files would be named differently and/or kept in different directories. If however, you want to to just use one .dll, and you want it to be available to 64-bit and 32-bit processes, then it must be compiled using the 32-bit csc.exe, and you must register it twice, once with the 32-bit RegAsm.exe and again with the 64-bit RegAsm.exe, changing from one .exe to the other using the SetBitness method.
'
'<a href="https://msdn.microsoft.com/en-us/library/78f4aasd.aspx">Command-line Building With csc.exe</a>
'
Class DotNetCompiler

    'Method RestartIfNotPrivileged
    'Remark: Elevates privileges if they are not already elevated. If app.GetUserInteractive is True, the user is first warned that the User Account Control dialog will open.
    Sub RestartIfNotPrivileged
        app.RestartIfNotPrivileged
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

    'Method SetExtension
    'Parameter: newExt
    'Remark: Sets the filename extension for the file to be compiled. Should be dll or exe. Optional. Default is dll.
    Sub SetExtension(newExt) : ext = newExt : End Sub

    'Method SetUserInteractive
    'Parameter: boolean
    'Remark: Sets userInteractive value. Setting to True can be useful for debugging. Default is True.
    Sub SetUserInteractive(newUserInteractive)
        app.SetUserInteractive newUserInteractive
    End Sub

    'Property GetUserInteractive
    'Returns: a boolean
    'Remark: Returns the userInteractive value.
    Property Get GetUserInteractive : GetUserInteractive = app.GetUserInteractive : End Property
    
    'Method SetOnUserCancelQuitApp
    'Parameter: boolean
    'Remark: Sets whether the calling app will quit if the user cancels out of a dialog.
    Sub SetOnUserCancelQuitApp(newOnUserCancelQuitApp)
        app.SetOnUserCancelQuitApp newOnUserCancelQuitApp
    End Sub

    'Property GetOnUserCancelQuitApp
    'Returns: boolean
    'Remark: Retrieves the boolean setting specifying whether the calling app will quit if the user cancels out of a dialog.
    Function GetOnUserCancelQuitApp
        GetOnUserCancelQuitApp = app.GetOnUserCancelQuitApp
    End Function

    'Method GenerateKeyPair
    'Remark: Generates a strong name key pair using Visual Studio's sn command. Requires Visual Studio to be installed. If the name for the .snk file is not specified, uses targetName.
    Sub GenerateKeyPair

        'validate

        ValidateName
        ValidateVisualStudio

        'build the arguments

        args = "/c cd """ & sourceFolder & """"
        args = args & " & """ & batFile & """"
        args = args & " & sn -k " & targetName & ".snk"
        If app.GetUserInteractive Then args = args & " & echo. & pause"

        'give an opt out

        msg = "Verify arguments for key generation"
        If app.GetUserInteractive Then If vbCancel = MsgBox(args, vbOKCancel, msg & " - " & scriptName) Then Quit

        'generate the strong name key pair

        sh.Run "%ComSpec% " & args, app.GetVisibility, synchronous

    End Sub

    Private Sub ValidateVisualStudio
        If fso.FileExists(batFile1) Then
            batFile = batFile1
        ElseIf fso.FileExists(batFile2) Then
            batFile = batFile2
        Else
            Err.Raise 1, scriptName, "Couldn't find either of the batch files hardcoded in """ & scriptName & """, which enable the Visual Studio strong name tool. If you don't have Visual Studio, but you have some version of the .NET framework in C:\Windows\Microsoft.NET\Framework, then you can still compile and register without a strong name: Remove or comment out the line in the .cs file with AssemblyKeyFileAttribute. You will receive a warning message when registering the .dll."
        End If
    End Sub

    Private Sub ValidateName
        If "" = targetName Then SetTargetName fso.GetBaseName(sourceFile)
        If "" = targetName Then Err.Raise 1, scriptName, scriptName & " needs a name. Use SetTargetName targetName, or pass in the name or the source file on the command line."
    End Sub

    Private Sub RequireFolder(folder, message)
        If Not fso.FolderExists(folder) Then Err.Raise 1, scriptName, message
    End Sub

    'Method Compile
    'Remark: Uses csc.exe to compile the .cs or .vb file specified by SetSourceFile or passed in on the command line.
    Sub Compile

        'validate

        If Not fso.FileExists(sourceFile) Then Err.Raise 1, scriptName, "Couldn't find the source file """ & sourceFile & """. Use SetSourceFile sourceFile, or pass in the source file on the command line."
        If Not "cs" = LCase(fso.GetExtensionName(sourceFile)) Then Err.Raise 1, scriptName, "A .cs file is required for compiling."
        RequireFolder exeFolder, "Couldn't find the .NET executables folder, " & L & exeFolder
        ValidateName

        'build the command arguments

        args = ""
        If "dll" = LCase(ext) Then args = args & " /target:library"
        If supressWarnings Then args = args & " /warn:0"
        If debug Then args = args & " /debug"

        'build the commmand string

        cmd = "%ComSpec% /c cd """ & sourceFolder & """"
        cmd = cmd & " & echo."
        cmd = cmd & " & """ & exeFolder & "\csc.exe"" /out:" & targetName & "." & ext & refs & args & " """ & sourceFile & """"
        cmd = cmd & " & echo. & echo OK to ignore warning CS1699 and BC41008. & echo."
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

        RequireFolder exeFolder, "Couldn't find the .NET executables folder, " & L & exeFolder
        If Not pc.Privileged Then Err.Raise 1, scriptName, "Registering a .dll requires elevated privileges."
        ValidateName

        'move the .dll to the target folder

        Dim dllFile : dllFile = targetName & ".dll"
        Dim sourceDllFile : sourceDllFile = sourceFolder & "\" & dllFile
        Dim targetDllFile : targetDllFile = targetFolder & "\" & dllFile
        If Not vfs.FoldersAreTheSame(sourceFolder, targetFolder) And Not unregistering Then
            If fso.FileExists(targetDllFile) Then
                vfs.DeleteFile targetDllFile
            End If
            fso.MoveFile sourceDllFile, targetDllFile
        End If

        'build the argument(s)

        args = "/c cd """ & targetFolder & """"
        args = args & " & echo. "
        args = args & " & """ & exeFolder & "\RegAsm.exe"""
        'args = args & " /tlb:" & targetName & ".tlb" 'create a type library
        args = args & " /codebase" 'if not putting .dll in the GAC
        args = args & " """ & dllFile & """"
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
        If 64 = bitness Then
            exeFolder = exeFolder64
        ElseIf 32 = bitness Then
            exeFolder = exeFolder32
        End If
    End Sub

    'Method Quit
    'Remark: Gracefully quits the hta/script, if allowed by settings.
    Sub Quit
        If app.GetOnUserCancelQuitApp Then 
            ReleaseObjectMemory
            app.Quit
        End If
    End Sub

    Private fso, sh, sa, vfs, pc, app
    Private batFile, batFile1, batFile2
    Private exeFolder, exeFolder64, exeFolder32
    Private sourceFile, sourceFolder, targetFolder, targetName, ext
    Private scriptName, thisFile
    Private refs, args, cmd
    Private supressWarnings, debug 'compiler settings
    Private unregistering 'registration setting
    Private msg, L
    Private synchronous, visible, hidden

    Sub Class_Initialize
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set sh = CreateObject("WScript.Shell")
        Set sa = CreateObject("Shell.Application")
        With CreateObject("includer")
            Execute(.read("VBSFileSystem"))
            Execute(.read("VBSApp"))
            Execute(.read("PrivilegeChecker"))
            Execute(.read("DotNetCompiler.config"))
        End With
        Set vfs = New VBSFileSystem
        Set app = New VBSApp
        Set pc = New PrivilegeChecker
        'get the filespec of the calling script
        thisFile = app.GetFullName
        SetSourceFile ""
        On Error Resume Next
            SetSourceFile app.GetArg(0)
        On Error Goto 0
        batFile1 = sh.ExpandEnvironmentStrings(batFile1) 'support environment variables in the .config file
        batFile2 = sh.ExpandEnvironmentStrings(batFile2)
        exeFolder32 = sh.ExpandEnvironmentStrings(exeFolder32)
        exeFolder64 = sh.ExpandEnvironmentStrings(exeFolder64)
        scriptName = fso.GetFileName(thisFile)
        L = vbLf & vbTab
        synchronous = True
        hidden = 0
        visible = 1
        unregistering = False 'defaults
        refs = ""
        app.SetOnUserCancelQuitApp True
        SetExtension "dll"
        SetSupressWarnings False
        SetDebug False
        SetBitness 64
        sh.CurrentDirectory = fso.GetParentFolderName(thisFile)
        SetTargetFolder sh.CurrentDirectory
        SetTargetName fso.GetBaseName(sourceFile)
        app.SetUserInteractive True
        'add references from the command-line, if any
        Dim args, nextArg : args = app.GetArgs
        If UBound(args) > 0 Then
            For i = 0 To UBound(args)
                nextArg = ""
                On Error Resume Next
                    nextArg = app.GetArg(i + 1)
                On Error Goto 0
                If "-ref" = LCase(app.GetArg(i)) Then If fso.FileExists(nextArg) Then AddRef nextArg
            Next
        End If
        Set args = Nothing

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
