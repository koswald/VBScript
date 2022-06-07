Dim vi 'Includer object
Set vi = New Includer

Function GetObj( className )
    Set GetObj = vi.GetObj( className )
End Function

Function LoadObject( className )
    Set LoadObject = vi.LoadObject( className )
End Function

Function Read( file )
    Read = vi.Read( file )
End Function

Function ReadFrom( relativePath, tempReferencePath )
    ReadFrom = vi.ReadFrom( relativePath, tempReferencePath )
End Function

Function LibraryPath
    LibraryPath = vi.LibraryPath
End Function

Sub SetLibraryPath( newPath )
    vi.SetLibraryPath newPath
End Sub