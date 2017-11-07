
Class VBSMessages

    Private L, item

    Private Sub Class_Initialize
        SetLineBreak(ErrRaiseLineBreak)
    End Sub

    Private Property Get file : file = "file" : End Property
    Private Property Get folder : folder = "folder" : End Property

    Private Sub SetItem(newItem) : item = newItem : End Sub
    Private Property Get ReqdArg : ReqdArg = "A command-line argument is required: " : End Property
    Property Get ReqdArgFile : SetItem(file) : ReqdArgFile = ReqdArg & "a " & item & "." & L & DragAndDrop : End Property
    Property Get ReqdArgFolder : SetItem(folder) : ReqdArgFolder = ReqdArg & "a " & item & "." & L & DragAndDrop : End Property
    Property Get ReqdArgPattern : ReqdArgPattern = ReqdArg & "a regex pattern." : End Property
    Private Property Get DragAndDrop : DragAndDrop = "You can drag the " & item & " onto this script or onto " & L & "a shortcut to this script, or place the shortcut in " & L & "the SendTo folder." : End Property

    Private Property Get ErrRaiseLineBreak : ErrRaiseLineBreak = vbLf & vbTab : End Property 'for Err.Raise descriptions (arg #3)
    Sub SetLineBreak(newLineBreak) : L = newLineBreak : End Sub

End Class
