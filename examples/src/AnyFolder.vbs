
'script for AnyFolder.hta, a SendTo drop target

Option Explicit

Const width = 350, height = 250 'window size: pixels
Const xPos = 80, yPos = 80 'window position: percent of screen

Sub Copy
    Transfer copyMode
    Cancel 'quit
End Sub

Sub Move
    Transfer moveMode
    Cancel 'quit
End Sub

Sub Transfer(mode)
    Dim i
    For i = 0 To UBound(items)
        TransferItem items(i), mode
    Next
End Sub

'copy or move a single file or folder
Sub TransferItem(sourceItem, mode)
    Dim targetItem : targetItem = targetFolder & "\" & fso.GetFileName(sourceItem)
    If sourceItem = targetItem Then Exit Sub 'if source and target are the same, don't transfer and especially don't delete!
    If fso.FolderExists(sourceItem) Then
        On Error Resume Next
            fso.CopyFolder sourceItem, targetItem, True
            If Err Then If vbCancel = MsgBox("Failed to copy folder """ & sourceItem & """ to """ & targetFolder & """.", vbInformation + vbOKCancel, app.GetFileName) Then app.Quit
        On Error Goto 0
    ElseIf fso.FileExists(sourceItem) Then
        On Error Resume Next
            fso.CopyFile sourceItem, targetItem, True
            If Err Then
                msg = Err.Description & vbLf & vbLf & "Failed to copy file """ & sourceItem & """ to """ & targetFolder & """."
                If vbCancel = MsgBox(msg, vbInformation + vbOKCancel, app.GetFileName) Then
                    app.Quit
                Else Exit Sub ' attempt to transfer the next item, if any
                End If
            End If
        On Error Goto 0
    End If
    If moveMode = mode Then
        DeleteSourceItem sourceItem, targetItem
    End If
    Dim msg
End Sub

'delete the source file or folder,
'as long as the copy was successful
Sub DeleteSourceItem(sourceItem, targetItem)
    If fso.FolderExists(targetItem) Then
        fso.DeleteFolder(sourceItem)
    ElseIf fso.FileExists(targetItem) Then
        fso.DeleteFile(sourceItem)
    End If
End Sub

'exit the html application
Sub Cancel
    app.Quit
End Sub

'event handler
Sub Document_OnKeyUp
    If EscKey = window.event.keyCode Then 
        Cancel
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

Const copyMode = 0, moveMode = 1
Const EscKey = 27 'window.event.keyCode for the Esc key
Const CKey = 67
Const MKey = 77
Dim sh, fso, app, hta, choose 'objects
Dim btnCopy, btnMove 'buttons
Dim targetFolder
Dim items 'array of files and/or folders

Sub Window_OnLoad
    InitializeWindow
    InstantiateObjects
    CreateHtmlElements
    ValidateArgs
    BrowseForFolder
End Sub

'initialize the window
Sub InitializeWindow
    Set hta = document.getElementsByTagName("application")(0)
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
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    With CreateObject("VBScripting.Includer")
        Execute .read("VBSApp")
        Set app = New VBSApp
    End With
    Set choose = CreateObject("VBScripting.FolderChooser")
    choose.Title = hta.applicationName & ":" & vbLf & " Browse to the target folder"
    choose.InitialDirectory = "%UserProfile%\z.{679F85CB-0220-4080-B29B-5540CC05AAB6}" 'see settings\+\Quick access folder.txt
End Sub

'create the HTML elements
Dim divText
Sub CreateHtmlElements

    'create the Copy button
    Set btnCopy = document.createElement("input")
    With btnCopy
        .type = "button"
        Set .onClick = GetRef("Copy")
        .value = "    Copy    "
        .style.margin = "1em"
        .style.marginLeft = "3em"
    End With
    document.body.insertBefore btnCopy

    'create the Move button
    Set btnMove = document.createElement("input")
    With btnMove
        .type = "button"
        Set .onClick = GetRef("Move")
        .value = "    Move    "
        .style.margin = "1em"
        .focus
    End With
    document.body.insertBefore btnMove

    'create a div for text
    Set divText = document.createElement("div")
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
        Cancel
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
            Cancel    
        End If
        s = s & items(i)
    Next
    divText.innerHtml = "<strong> Ready to send: </strong>" & s
End Sub

'prompt the user to choose a target folder
Sub BrowseForFolder
    DisableCopyAndMoveButtons True
    targetFolder = choose.FolderName
    If Not fso.FolderExists(targetFolder) Then Cancel
    DisableCopyAndMoveButtons False
    btnMove.focus
    divText.innerHtml = divText.innerHtml & "<br /><strong> To: </strong><br />" & targetFolder
End Sub
