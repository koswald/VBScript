'Test FileChooser.dll

Option Explicit : Initialize

With New TestingFramework

    .Describe "VBScripting.FileChooser object"
        Dim fc
        Set fc = CreateObject( "VBScripting.FileChooser" )

    .It "should have a Title property"
        On Error Resume Next
            x = fc.Title
            fc.Title = x
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should have a Multiselect property"
        On Error Resume Next
            x = fc.Multiselect
            fc.Multiselect = x
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should have a DereferenceLinks property"
        On Error Resume Next
            x = fc.DereferenceLinks
            fc.DereferenceLinks = x
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should have a DefaultExt property"
        On Error Resume Next
            x = fc.DefaultExt
            fc.DefaultExt = x
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should have a Filter property"
        On Error Resume Next
            x = fc.Filter
            fc.Filter = x
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should have a FilterIndex property"
        On Error Resume Next
            x = fc.FilterIndex
            fc.FilterIndex = x
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should have an InitialDirectory property"
        On Error Resume Next
            x = fc.InitialDirectory
            fc.InitialDirectory = x
            .AssertEqual Err.Description, ""
        On Error Goto 0
    .It "should open at the current directory"
        .AssertEqual fc.InitialDirectory, sh.CurrentDirectory

    .It "should have an ERInitialDirectory property (get only)"
        On Error Resume Next
            x = fc.ERInitialDirectory
            .AssertEqual Err.Description, ""
        On Error Goto 0
        .ShowPendingResult
        WScript.StdOut.WriteLine "          " & _
            "(ER=Expanded, Resolved)"

    .It "should resolve a relative path"
        fc.InitialDirectory = ".."
        .AssertEqual fc.ERInitialDirectory, fso.GetAbsolutePathName( ".." )

    .It "should expand an environment variable"
        fc.InitialDirectory = "%UserProfile%"
        .AssertEqual fc.ERInitialDirectory, sh.ExpandEnvironmentStrings("%UserProfile%")

End With

Cleanup

Sub Cleanup
    Set fc = Nothing
    Set sh = Nothing
    Set fso = Nothing
    Set includer = Nothing
End Sub

Dim sh, fso, includer
Dim x
Sub Initialize
    Set sh = CreateObject( "WScript.Shell" )
    Set fso = CreateObject( "Scripting.FileSystemObject" )
    Set includer = CreateObject( "VBScripting.Includer" )
    ExecuteGlobal includer.Read( "TestingFramework" )
End Sub


