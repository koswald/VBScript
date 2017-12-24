
'fixture for FolderChooserX.spec.vbs

Set args = WScript.Arguments
progId = args.item(0)

'Set the properties and get the folder path (FolderName)
With CreateObject(progId)
    If Not "." = args.item(1) Then
        .InitialDirectory = args.item(1)
    End If
    If Not "." = args.item(2) Then
        .Title = args.item(2)
    End If
    folder = .FolderName
End With

'Save the folder path to file
With CreateObject("Scripting.FileSystemObject")
    Const ForWriting = 2
    Const CreateNew = True
    Set stream = .OpenTextFile( _
        "fixture\FolderChooserX.fixture.txt", _
        ForWriting, CreateNew)
    stream.WriteLine folder
    stream.Close
    Set stream = Nothing
End With
