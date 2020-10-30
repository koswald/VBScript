' Copies all the files and folders in the specified SourceFolder to the specified TargetFolder.

' Running the script with elevated privileges is not required, provided that no new folders need to be created in a target folder that requires special permissions. 

' For example, if the target folder is a non-existing subfolder in %ProgramFiles%, then the script must be run with elevated privileges, or else the target folder must be created in advance. The CopyHere method itself does not require that the script be run elevated, even when the target is in the Program Files folder, because it will ask for permissions as necessary, without resorting to the UAC dialog; however, the CopyHere method does require that the target folder exists.

With New FolderCopier
    .SourceFolder = .Parent(WScript.ScriptFullName)
    .TargetFolder = "%ProgramFiles%\KOswald"
    .Copy
End With

Class FolderCopier

    Sub Copy
        EnsureInitialized
        With CreateObject("Shell.Application")
            .Namespace(TargetFolder).CopyHere SourceFolder
        End With
    End Sub

    Sub EnsureInitialized
        If IsEmpty(sourceFolder_) _
        Or IsEmpty(targetFolder_) Then
            Err.Raise 1,, "You must specify the SourceFolder and TargetFolder properties."
        End If
    End Sub

    Private sourceFolder_
    Public Property Let SourceFolder(newValue)
        sourceFolder_ = Resolve(Expand(newValue))
        If Not fso.FolderExists(sourceFolder_) Then
            Err.Raise 3,, "Cannot find the source folder '" & sourceFolder_ & "'"
        End If
    End Property
    Public Property Get SourceFolder
        SourceFolder = sourceFolder_
    End Property

    Private targetFolder_
    Public Property Let TargetFolder(newValue)
        targetFolder_ = Resolve(Expand(newValue))
        MakeFolder targetFolder_
    End Property
    Public Property Get TargetFolder
        TargetFolder = targetFolder_
    End Property

    Function MakeFolder(sFolder)
        If "" = sFolder Then Err.Raise 2, "FolderCopier.MakeFolder", "No folder specified."
        If Not fso.FolderExists(Parent(Expand(sFolder))) Then
    	    MakeFolder(Parent(sFolder))	'Recurse: create parent before child
        End If
        If Not fso.FolderExists(Expand(sFolder)) Then
            On Error Resume Next
                fso.CreateFolder(Expand(sFolder)) 
                If Err Then MsgBox Err.Description & "." & vbLf & "Error creating folder " & vbLf & sFolder & vbLf & vbLf & "Either create the folder and try again, or else run this script with elevated privileges.", vbInformation + vbSystemModal, WScript.ScriptName : WScript.Quit
            On Error Goto 0
        End If
        If fso.FolderExists(Expand(sFolder)) Then
            MakeFolder = True
        Else MakeFolder = False 'folder could not be created
        End If
    End Function

    Function Parent(child)
        Parent = fso.GetParentFolderName(child)
    End Function

    Function Expand(str)
        Expand = sh.ExpandEnvironmentStrings(str)
    End Function

    Function Resolve(path)
        Resolve = fso.GetAbsolutePathName(path)
    End Function

    Private fso
    Private sh

    Sub Class_Initialize
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set sh = CreateObject("WScript.Shell")
    End Sub

End Class
