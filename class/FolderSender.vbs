'The FolderSender class supplies methods that copy or move (send) the specified SourceFolder to the specified TargetFolder. Operator action may be  required. 

Class FolderSender

    Private sa 'Shell.Application object
    Private fso 'Scripting.FileSystemObject 
    Private sh 'WScript.Shell object
    Private app 'VBSApp object
    Private elevateMsg 'string 
    
    Sub Class_Initialize
        Set fso = CreateObject( "Scripting.FileSystemObject" )
        Set sh = CreateObject( "WScript.Shell" )
        Set sa = CreateObject( "Shell.Application" )
        elevateMsg = "Click OK to elevate privileges. The User Account Control dialog will appear and the script will restart."
        With CreateObject( "VBScripting.Includer")
            Execute .Read( "VBSApp" )
        End With
        Set app = New VBSApp
    End Sub

    'Method Copy
    'Remarks: Copies a folder. The SourceFolder and TargetFolder properties must be specified in advance or else an error will occur. A familiar Windows-native graphical interface appears for sizeable operations or when it is necessary to overwrite existing files or to elevate privileges: Operator action may be required. 
    Sub Copy
        EnsureInitialized
        sa.Namespace(TargetFolder).CopyHere SourceFolder
    End Sub
    
    'Method Move
    'Remarks: Moves a folder. The SourceFolder and TargetFolder properties must be specified in advance or else an error will occur. A familiar Windows-native graphical interface appears for sizeable operations or when it is necessary to overwrite existing files or to elevate privileges: Operator action may be required. 
    Sub Move
        EnsureInitialized
        sa.Namespace(TargetFolder).MoveHere SourceFolder
    End Sub

    Sub EnsureInitialized
        If IsEmpty(sourceFolder_) _
        Or IsEmpty(targetFolder_) Then
            Err.Raise 500,, "The SourceFolder and TargetFolder properties must be specified."
        End If
    End Sub

    'Property SourceFolder
    'Parameter: a string (folder)
    'Returns: a string (folder)
    'Remarks: Required. Sets or gets the source folder for the Copy and Move methods. Relative paths are allowed. Environment variables are allowed. The source folder must exist or an error will occur.
    Public Property Let SourceFolder(newValue)
        sourceFolder_ = Resolve(Expand(newValue))
        If Not fso.FolderExists(sourceFolder_) Then
            Err.Raise 450,, "Cannot find the source folder '" & sourceFolder_ & "'"
        End If
    End Property
    Public Property Get SourceFolder
        SourceFolder = sourceFolder_
    End Property
    Private sourceFolder_

    'Property TargetFolder
    'Parameter: a string (folder)
    'Returns: a string (folder)
    'Remarks: Required. Sets or gets the target folder for the Copy and Move methods. Relative paths are allowed (see the CurrentDirectory property). Environment variables are allowed. The target folder will be created if it does not exist. The User Account Control dialog may appear to request permission to create a folder if it is in a location that has restricted write permissions such as %ProgramFiles%.
    Public Property Let TargetFolder(newValue)
        targetFolder_ = Resolve(Expand(newValue))
        MakeFolder targetFolder_
    End Property
    Public Property Get TargetFolder
        TargetFolder = targetFolder_
    End Property
    Private targetFolder_

    'Property CurrentDirectory
    'Parameter: a string (folder)
    'Returns a string (folder)
    'Remarks: Gets or sets the current directory or working directory. Relative paths are allowed. Environment variables are allowed.
    Property Let CurrentDirectory( ByVal newValue )
        newValue = Resolve( Expand( newValue ))
        If Not fso.FolderExists( newValue ) Then
            Err.Raise 505,, "Couldn't find the folder '" & newValue & "'."
        End If
        sh.CurrentDirectory = newValue
    End Property
    Property Get CurrentDirectory
        CurrentDirectory = sh.CurrentDirectory
    End Property

    Function MakeFolder( ByVal sFolder )
        Dim m, j, s 'MsgBox arguments
        Dim errored 'boolean
        Dim errDesc 'string
        Dim errNum 'integer
        sFolder = Expand( sFolder )
        
        'validate
        If "" = sFolder Then
            Err.Raise 450, "FolderSender.MakeFolder", "No folder specified."
        End If
        
        'recurse: create parent before child
        If Not fso.FolderExists(Parent(sFolder)) Then
    	    MakeFolder(Parent(sFolder))
        End If

        'attempt to create the folder
        errored = False
        If Not fso.FolderExists(sFolder) Then
            On Error Resume Next
                fso.CreateFolder(sFolder)
                If Err Then
                    errored = True
                    errDesc = Err.Description
                    errNum = Err.Number
                End If
            On Error Goto 0
        End If
        If errored Then

            'provide opt out for elevating permissions
            m = "Error creating folder" & vbLf
            m = m & sFolder & vbLf & vbLf
            m = m & errDesc & " ( &H"
            m = m & Hex(errNum) & " )." & vbLf & vbLf 
            m = m & elevateMsg
            j = vbInformation + vbSystemModal
            j = j + vbOkCancel + vbDefaultButton2
            If IsEmpty( app ) Then
                s = ""
            Else s = app.GetFileName
            End If
            If vbCancel = MsgBox(m, j, s) Then
                On Error Resume Next
                    WScript.Quit
                    Self.Close
                On Error Goto 0
            End If

            'elevate privileges
            If IsEmpty( app ) Then
                'before the project Setup.vbs has been run, the VBSApp object is not available: assume that the calling script is not an .hta; restart the calling script with elevated privileges without retaining command-line arguments
                sa.ShellExecute "wscript", """" & WScript.ScriptFullName & """",, "runas"
                WScript.Quit
            Else
                'restart the calling script/hta,
                'retaining command-line arguments
                With app
                    .SetUserInteractive False
                    .RestartUsing .GetExe, .DoExit, .DoElevate
                End With
            End If
        End If
        
        'set return value
        If fso.FolderExists(sFolder) Then
            MakeFolder = True
        Else MakeFolder = False
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

    'Testability Properties
    
    Property Let MockSourceFolder( newValue )
        sourceFolder_ = newValue
    End Property
    
    Property Let MockTargetFolder( newValue )
        targetFolder_ = newValue
    End Property
End Class
