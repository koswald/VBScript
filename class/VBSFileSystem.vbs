
Class VBSFileSystem

    Private oVBSNatives, oVBSEnvironment
    Private referencePath, savedCurrentDirectory, savedReferencePath

    Private Sub Class_Initialize 'event fires on object instantiation
        With CreateObject("includer") : On Error Resume Next
            ExecuteGlobal(.read("VBSNatives"))
            ExecuteGlobal(.read("VBSEnvironment"))
        End With : On Error Goto 0
        Set oVBSNatives = New VBSNatives
        Set oVBSEnvironment = New VBSEnvironment

        SetReferencePath(defaultReferencePath)
        SaveReferencePath
    End Sub

    Property Get n : Set n = oVBSNatives : End Property
    Property Get natives : Set natives = n : End Property
    Property Get shell : Set shell = sh : End Property
    Property Get sh : Set sh = n.sh : End Property
    Property Get fso : Set fso = n.fso : End Property
    Property Get args : Set args = a : End Property
    Property Get a : Set a = n.a : End Property

    Property Get SName : SName = WScript.ScriptName : End Property 'script name; i.e. the name of the calling script
    Property Get SFullName : SFullName = WScript.ScriptFullName : End Property 'script filespec (with path)
    Property Get SBaseName : SBaseName = fso.GetBaseName(SName) : End Property 'script name without filename extension
    Property Get SFolderName : SFolderName = Parent(SFullName) : End Property 'script's folder

    Property Get m : Set m = oVBSMessages : End Property
    Property Get msgs : Set msgs = m : End Property

    'Function MakeFolder
    'Parameter: a path
    'Returns False if the folder could not be created.
    'Remark Create a folder, and if necessary create also its parent, grandparent, etc.
    Function MakeFolder(sFolder)
        MakeFolder = True

	    If Not fso.FolderExists(Parent(Expand(sFolder))) Then
		    MakeFolder(Parent(sFolder))	'Recurse: create parent before child
	    End If
	    If Not fso.FolderExists(Expand(sFolder)) Then fso.CreateFolder(Expand(sFolder)) 'create folder
	    If Not fso.FolderExists(Expand(sFolder)) Then MakeFolder = False 'folder could not be created
    End Function

    'Property Parent
    'Parameter: a string representing a folder or file or registry key
    'Returns the parent of the folder or file or registry key, or removes a trailing backslash
    'Remark: The parent need not exist.
    Function Parent(string)
        If 0 = InStr(string, "\") Then Parent = "" : Exit Function
        Parent = Left(string, InStrRev(string, "\") - 1)
    End Function

    Private Sub SaveCurrentDirectory : savedCurrentDirectory = sh.CurrentDirectory : End Sub
    Private Sub RestoreCurrentDirectory : sh.CurrentDirectory = savedCurrentDirectory : End Sub
    Private Property Get defaultReferencePath : defaultReferencePath = Parent(WScript.ScriptFullName) : End Property
    Private Sub SaveReferencePath : savedReferencePath = referencePath : End Sub
    Private Sub RestoreReferencePath : referencePath = savedReferencePath : End Sub

    'Method SetReferencePath
    'Parameter: a path
    'Remark: Call this method, if desired, before calling the property Resolve in order specify the base path against which relative paths should be referenced from. By default, the reference path is the parent folder of the calling script.
    Sub SetReferencePath(newPath) : referencePath = newPath : End Sub
    Property Get GetReferencePath : GetReferencePath = referencePath : End Property

    'Property Resolve
    'Returns a resolved path
    'Parameter: a relative path
    'Remark: Resolves a relative path (e.g. "../lib/WMI.vbs"), to an absolute path (e.g. "C:\Users\user42\lib\WMI.vbs"). The relative path is by default relative to the parent folder of the calling script, but this behavior can be changes with SetReferencePath. See also property ResolveTo.
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
        n.sa.ShellExecute fs.Expand(cmd), fs.Expand(args_), fs.Expand(workingFolder), "runas"
    End Sub

End Class
