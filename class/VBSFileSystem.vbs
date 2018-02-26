
'General utility functions
'
Class VBSFileSystem

    Private env
    Private msgs
    Private sh, fso
    Private referencePath, savedCurrentDirectory, savedReferencePath
    Private scriptName, scriptFullName
    Private forceDelete

    Private Sub Class_Initialize 'event fires on object instantiation
        With CreateObject("VBScripting.Includer")
            Execute .read("VBSMessages")
            Execute .read("VBSEnvironment")
        End With

        Set msgs = New VBSMessages
        Set env = New VBSEnvironment
        Set sh = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")

        On Error Resume Next
            scriptFullName = WScript.ScriptFullName
            If Err Then scriptFullName = Replace(Replace(document.location.href, "file:///", ""), "%20", "") '.hta file
        On Error Goto 0

        scriptName = fso.GetFileName(scriptFullName)
        SetReferencePath defaultReferencePath
        SaveReferencePath
        SetForceDelete False
    End Sub

    'Property SBaseName
    'Returns a file name, no extension
    'Remark: Returns the name of the calling script, without the file name extension.
    Property Get SBaseName : SBaseName = fso.GetBaseName(SName) : End Property 'script name without filename extension

    'Property SName
    'Returns a file name
    'Remark: Returns the name of the calling script, including file name extension
    Property Get SName : SName = scriptName : End Property 'script name; i.e. the name of the calling script

    'Property SFullName
    'Returns a filespec
    'Remark: Returns the filespec of the calling script
    Property Get SFullName : SFullName = scriptFullName : End Property 'script filespec (with path)

    'Property SFolderName
    'Returns a folder
    'Remark: Returns the parent folder of the calling script.
    Property Get SFolderName : SFolderName = Parent(SFullName) : End Property 'script's folder

    'Function MakeFolder
    'Parameter: a path
    'Returns a boolean
    'Remark: Create a folder, and if necessary create also its parent, grandparent, etc. Returns False if the folder could not be created.
    Function MakeFolder(sFolder)
        If "" = sFolder Then Err.Raise 1, "VBSFileSystem.MakeFolder", "No folder specified."
        If Not fso.FolderExists(Parent(Expand(sFolder))) Then
    	    MakeFolder(Parent(sFolder))	'Recurse: create parent before child
        End If
        If Not fso.FolderExists(Expand(sFolder)) Then fso.CreateFolder(Expand(sFolder)) 'create folder
        If fso.FolderExists(Expand(sFolder)) Then
            MakeFolder = True
        Else MakeFolder = False 'folder could not be created
        End If
    End Function

    'Property Parent
    'Parameter: a folder, file, or registry key
    'Returns the item's parent
    'Remark: Returns the parent of the folder or file or registry key, or removes a trailing backslash. The parent need not exist.
    Function Parent(string)
        If 0 = InStr(string, "\") Then Parent = "" : Exit Function
        Parent = Left(string, InStrRev(string, "\") - 1)
    End Function

    'Method SetReferencePath
    'Parameter: a path
    'Remark: Optional. Specifies the base path from which relative paths should be referenced. By default, the reference path is the parent folder of the calling script. See also Resolve and ResolveTo.
    Sub SetReferencePath(newPath) : referencePath = newPath : End Sub

    Private Property Get defaultReferencePath : defaultReferencePath = SFolderName : End Property
    Property Get GetReferencePath : GetReferencePath = referencePath : End Property
    Private Sub SaveCurrentDirectory : savedCurrentDirectory = sh.CurrentDirectory : End Sub
    Private Sub RestoreCurrentDirectory : sh.CurrentDirectory = savedCurrentDirectory : End Sub
    Private Sub SaveReferencePath : savedReferencePath = referencePath : End Sub
    Private Sub RestoreReferencePath : referencePath = savedReferencePath : End Sub

    'Property Resolve
    'Returns a resolved path
    'Parameter: a relative path
    'Remark: Resolves a relative path (e.g. "../lib/WMI.vbs"), to an absolute path (e.g. "C:\Users\user42\lib\WMI.vbs"). The relative path is by default relative to the parent folder of the calling script, but this behavior can be changed with SetReferencePath. See also property ResolveTo.
    Function Resolve(path)
        SaveCurrentDirectory
        sh.CurrentDirectory = referencePath 'in case the path is relative, set the reference folder for .GetAbsolutePathName
        Resolve = fso.GetAbsolutePathName(Expand(path))
        RestoreCurrentDirectory
    End Function

    'Property ResolveTo
    'Returns a resolved path
    'Parameter: relativePath, absolutePath
    'Remark: Resolves the specified relative path, e.g. "../lib/WMI.vbs", relative to the specified absolute path, and returns the resolved absolute path, e.g. "C:\Users\user42\lib\WMI.vbs". Environment variables are allowed.
    Function ResolveTo(relativePath, absolutePath)
        SaveReferencePath
        SetReferencePath Expand(absolutePath) 'in case the path is relative, set the reference folder for .GetAbsolutePathName
        ResolveTo = Resolve(relativePath)
        RestoreReferencePath
    End Function

    'Property Expand
    'Returns an expanded string
    'Parameter: a string
    'Remark: Expands environment strings. E.g. %WinDir% => C:\Windows
    Property Get Expand(str) : Expand = sh.ExpandEnvironmentStrings(str) : End Property

    'Method Elevate
    'Parameters: command, arguments, folder
    'Remarks: Runs the specified command with elevated privileges, with the specified arguments and working folder
    Sub Elevate(cmd, args_, workingFolder)
        sa.ShellExecute fs.Expand(cmd), fs.Expand(args_), fs.Expand(workingFolder), "runas"
    End Sub

    'Property FoldersAreTheSame
    'Parameters: folder1, folder2
    'Returns: a boolean
    'Remark: Determines whether the two specified folders are the same. If so, returns True.
    Property Get FoldersAreTheSame(folder1, folder2)
        FoldersAreTheSame = CBool(LCase(Expand(folder1)) = LCase(Expand(folder2)))
    End Property

    'Method DeleteFile
    'Parameter: filespec
    'Remark: Deletes the specified file.
    Sub DeleteFile(filespec)
        fso.DeleteFile filespec, forceDelete
    End Sub

    'Method SetForceDelete
    'Parameter: boolean
    'Remark: Controls the behavior of the DeleteFile method: Specify True to force a file deletion. Optional. Default is False.
    Sub SetForceDelete(newForceDelete)
        forceDelete = newForceDelete
    End Sub

End Class
