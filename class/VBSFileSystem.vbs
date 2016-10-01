
Class VBSFileSystem

    Private oVBSNatives, oVBSEnvironment
    Private relativePath, savedCurrentDirectory, savedRelativePath

    Private Sub Class_Initialize 'event fires on object instantiation
        With CreateObject("includer") : On Error Resume Next
            ExecuteGlobal(.read("VBSNatives"))
            ExecuteGlobal(.read("VBSEnvironment"))
        End With : On Error Goto 0
        Set oVBSNatives = New VBSNatives
        Set oVBSEnvironment = New VBSEnvironment

        SetRelativePath(defaultRelativePath)
        SaveRelativePath
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

    Property Get env : Set env = oVBSEnvironment : End Property

    Property Get m : Set m = oVBSMessages : End Property
    Property Get msgs : Set msgs = m : End Property

    'Function MakeFolder
    'Parameter: a path
    'Returns False if the folder could not be created.
    'Remark Create a folder, and its parent, grandparent, etc.
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
    Private Property Get defaultRelativePath : defaultRelativePath = Parent(WScript.ScriptFullName) : End Property

    'Method SetRelativePath
    'Parameter: a path
    'Remark: Call this method, if desired, before calling the property Resolve in order specify the base path against which relative paths should be referenced from. By default, the reference path is the parent folder of the calling script.
    Sub SetRelativePath(newPath) : relativePath = newPath : End Sub
    Property Get GetRelativePath : GetRelativePath = relativePath : End Property

    '''Method SaveRelativePath
    '''Remark: Save the current relative path in order to be restored later with RestoreRelativePath
    Private Sub SaveRelativePath
        savedRelativePath = relativePath
    End Sub

    '''Method RestoreRelativePath
    '''Remark: Restore the relative path to the saved value or its initial value if it wasn't saved.
    Private Sub RestoreRelativePath
        relativePath = savedRelativePath
    End Sub

    'Property Resolve
    'Returns a resolved path
    'Parameter: a relative path
    'Remark: Resolves a relative path (e.g. "../lib/WMI.vbs"), to an absolute path (e.g. "C:\Users\user42\lib\WMI.vbs"). The relative path is by default relative to the parent folder of the calling script, but can aslo be set with SetRelativePath. See also property ResolveTo.
    Function Resolve(path)
        SaveCurrentDirectory
        sh.CurrentDirectory = relativePath 'in case the path is relative, set the reference folder for .GetAbsolutePathName
        Resolve = fso.GetAbsolutePathName(Expand(path))
        RestoreCurrentDirectory
    End Function

    'Property ResolveTo
    'Returns a resolved path
    'Parameter: relativePath, absolutePath
    'Remark: Resolves the specified relative path (e.g. "../lib/WMI.vbs"), relative to the specified absolute path to another absolute path (e.g. "C:\Users\user42\lib\WMI.vbs")
    Function ResolveTo(aRelativePath, anAbsolutePath)
        SaveRelativePath
        SetRelativePath Expand(anAbsolutePath) 'in case the path is relative, set the reference folder for .GetAbsolutePathName
        ResolveTo = Resolve(aRelativePath)
        RestoreRelativePath
    End Function

    'Property Expand
    'Returns an expanded string
    'Parameter: a string
    'Remark: Expands environment strings. E.g. %WinDir% => C:\Windows
    Property Get Expand(str) : Expand = sh.ExpandEnvironmentStrings(str) : End Property

End Class
