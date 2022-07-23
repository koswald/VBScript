'Script for AnyFolder.hta, a SendTo drop target

'A UI will be presented, showing the items to be copied or moved, read from the command line. The Copy and Move buttons are enabled after the user selects the target folder.

Option Explicit
Const copyMode = 0, moveMode = 1
Const EscKey = 27 'window.event.keyCode for the Esc key
Const CKey = 67
Const MKey = 77
Const width = 350, height = 250 'window size: pixels
Const xPos = 80, yPos = 80 'window position: percent of screen

Dim sh, fso, sa, hta 'native Windows objects
Dim app, choose 'project objects
Dim btnCopy, btnMove 'buttons
Dim divText 'html element
Dim targetFolder 'string: folder path
Dim items 'array of strings: filespecs and/or folder paths

Sub Window_OnLoad
    InitializeWindow
    InstantiateObjects
    CreateHtmlElements
    ValidateArgs
    BrowseForFolder
End Sub

'initialize the window
Sub InitializeWindow
    Set hta = document.getElementsByTagName( "application" )(0)
    self.ResizeTo width, height
    With document.parentWindow.screen
        self.MoveTo _
            (.availWidth - width) * xPos * .01, _
            (.availHeight - height) * yPos * .01
    End With

    document.title = hta.applicationName 'title bar
    document.body.style.whitespace = "nowrap"
End Sub

'instantiate objects
Sub InstantiateObjects
    Set sh = CreateObject( "WScript.Shell" )
    Set fso = CreateObject( "Scripting.FileSystemObject" )
    Set sa = CreateObject( "Shell.Application" )
    With CreateObject( "VBScripting.Includer" )
        Execute .Read( "VBSApp" )
        Set app = New VBSApp
    End With
    Set choose = CreateObject( "VBScripting.FolderChooser" )
    choose.Title = hta.applicationName & ":" & vbLf & " Browse to the target folder"
    choose.InitialDirectory = "%UserProfile%"
End Sub

'create the HTML elements
Sub CreateHtmlElements

    'create the Cop]y button
    Set btnCopy = document.createElement( "input" )
    With btnCopy
        .type = "button"
        Set .onClick = GetRef( "Copy" )
        .value = "    Copy    "
        .style.margin = "1em"
        .style.marginLeft = "3em"
    End With
    document.body.insertBefore btnCopy

    'create the Move button
    Set btnMove = document.createElement( "input" )
    With btnMove
        .type = "button"
        Set .onClick = GetRef( "Move" )
        .value = "    Move    "
        .style.margin = "1em"
        .focus
    End With
    document.body.insertBefore btnMove

    'create a div for text
    Set divText = document.createElement( "div" )
    With divText
        .style.fontFamily = "sans-serif"
        .style.fontSize = "75%"
    End With
    document.body.insertBefore divText

End Sub

'validate the command-line arguments
Sub ValidateArgs
    items = app.GetArgs

    'require at least one argument
    If -1 = UBound(items) Then
        MsgBox "Argument(s) required.", vbExclamation, hta.applicationName
        Self.Close
    End If

    'require all arguments to be an existing file or folder
    Dim i, s : s = ""
    For i = 0 To UBound(items)
        If fso.FileExists(items(i)) Then
            s = s & "<br />" & "File:  "
        ElseIf fso.FolderExists(items(i)) Then
            s = s & "<br />" & "Folder: "
        Else
            MsgBox """" & items(i) & """ is not an existing file or folder.", vbExclamation, hta.applicationName
            Self.Close
        End If
        s = s & items(i)
    Next
    divText.innerHtml = "<strong> Ready to send: </strong>" & s
End Sub

'prompt the user to choose a target folder
Sub BrowseForFolder
    DisableCopyAndMoveButtons True
    targetFolder = choose.FolderName
    If Not fso.FolderExists(targetFolder) Then 
        Self.Close 'user cancelled
    End If
    DisableCopyAndMoveButtons False
    btnMove.focus
    divText.innerHtml = divText.innerHtml & "<br /><strong> To: </strong><br />" & targetFolder
End Sub

Sub Copy
    Transfer copyMode
    Self.Close 'quit
End Sub

Sub Move
    Transfer moveMode
    Self.Close 'quit
End Sub

Sub Transfer(mode)
    Dim i
    For i = 0 To UBound(items)
        TransferItem items(i), mode
    Next
End Sub

'copy or move a single file or folder
Sub TransferItem(sourceItem, mode)
    Dim targetItem
    targetItem = targetFolder & "\" & fso.GetFileName(sourceItem)
    If LCase(sourceItem) = LCase(targetItem) Then 
        'if source and target are the same, don't transfer
        Exit Sub
    End If
    If Not fso.FolderExists(sourceItem) _
    And Not fso.FileExists(sourceItem) Then
        'Most likely, the command-line couldn't hold all of the desired arguments, and the current item was cut off. Attempt to continue anyway, with the next file/command-line argument. Only valid items will be transferred.
        Exit Sub
    End If
    If copyMode = mode Then
        sa.Namespace(targetFolder).CopyHere sourceItem
    ElseIf moveMode = mode Then
        sa.Namespace(targetFolder).MoveHere sourceItem
    Else Err.Raise 5,, "Mode specified incorrectly:  it must be either copyMode or moveMode."
    End If
End Sub

'event handler
Sub Document_OnKeyUp
    If EscKey = window.event.keyCode Then
        Self.Close
    ElseIf MKey = window.event.keyCode Then
        Move
    ElseIf CKey = window.event.keyCode Then
        Copy
    End If
End Sub

'disable or enable the buttons
Sub DisableCopyAndMoveButtons(newValue)
    btnCopy.disabled = newValue
    btnMove.disabled = newValue
End Sub
