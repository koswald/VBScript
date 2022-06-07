' Test the SetupHelper object

' This test can be run even when the project's Windows Script Components are not registered, provided that it is located in one of the child (not grandchild) folders of the project folder, in order that the class files can be located.

Option Explicit
Dim suh ' SetupHelper object under test
Dim fso ' Scripting.FileSystemObject
Dim sh ' WScript.Shell object
Dim desc ' error description
Dim errNumber
Dim expected, actual
Const guidPattern = "^[A-Fa-f\d]{8}-[A-Fa-f\d]{4}-[A-Fa-f\d]{4}-[A-Fa-f\d]{4}-[A-Fa-f\d]{12}$"

' Ensure that cscript.exe is the host

If "wscript.exe" = LCase(Right(WScript.FullName,11)) Then
    MsgBox WScript.ScriptName & " syntax:" & vbLf & vbLf & _
        "cscript [/nologo] .\" & WScript.ScriptName, _
        vbInformation, WScript.ScriptName
    WScript.Quit
End If

Set fso = CreateObject( "Scripting.FileSystemObject" )
Set sh = CreateObject( "WScript.Shell" )

Execute fso.OpenTextFile("../class/TestingFramework.vbs").ReadAll
With New TestingFramework

    .Describe "SetupHelper class"
        Execute fso.OpenTextFile("../class/SetupHelper.vbs").ReadAll
        Set suh = New SetupHelper

    ' If either of the properties ComponentFolder and ConfigFile are not explicitly set before the Setup method is called, then these properties will be set with the defaults when .Setup is called. These defaults should cause an error if the calling script is not in the project root folder.

    .It "should err: calling script not in root; .Setup called before .ComponentFolder"
        suh.ConfigFile = "fixture/SetupHelper.config"
        On Error Resume Next
            suh.Init
            errNumber = Err.Number
        On Error Goto 0
        .AssertEqual errNumber, -2146697211

    .It "should err: calling script not in root; .Setup called before .ConfigFile"
        Set suh = Nothing
        Set suh = New SetupHelper
        suh.ComponentFolder = "../class/wsc"
        On Error Resume Next
            suh.Init
            desc = Err.Description
        On Error Goto 0
        .AssertEqual desc, "File not found"

    .It "should instantiate the string formatter"
        actual = suh.Format( Array("the %s", "fox" ))
        expected = "the fox"
        .AssertEqual actual, expected

    .It "should read the .config file"
        suh.ConfigFile = "fixture/SetupHelper.config"
        actual = UBound(suh.WscGuids)
        expected = 18
        .AssertEqual actual, expected

    .It "should identify an interface progid (partial)"
        actual = suh.Char2IsUpperCase( "IFileChooser" )
        expected = True
        .AssertEqual actual, expected

    .It "should identify a class progig (partial)"
        actual = suh.Char2IsUpperCase( "FileChooser" )
        expected = False
        .AssertEqual actual, expected

    .It "should not err on invalid interface guid"
        suh.ConfigFile = "fixture\SetupHelper2.config"
        On Error Resume Next
            suh.EnsureValidRegData _
                suh.DllGuids, 2, 2, -1, guidPattern
            desc = Err.Description
            errNumber = Err.Number
        On Error Goto 0
        actual = desc & errNumber
        expected = "0"
        .AssertEqual actual, expected

    .It "should err on invalid class guid"
        suh.ConfigFile = "fixture\SetupHelper3.config"
        ' WScript.StdOut.WriteLine UBound(suh.DllGuids)
        ' .ShowPendingResult
        On Error Resume Next
            suh.EnsureValidRegData _
                suh.DllGuids, 2, 2, -1, guidPattern
            desc = Err.Description
        On Error Goto 0
        actual = Left(desc, 26)
        expected = "Invalid registration data:"
        .AssertEqual actual, expected

    .It "should indicate a key exists"
        actual = suh.KeyExists( suh.HKCU, "Software\Microsoft" )
        expected = True
        .AssertEqual actual, expected

    .It "should indicate a key doesn't exist"
        actual = suh.KeyExists( suh.HKCU, "Software\Luddite Software" )
        expected = False
        .AssertEqual actual, expected
    .ShowPendingResult

    .It "should indicate a root-level key exists"
        actual = suh.KeyExists( suh.HKCU, "Software" )
        expected = True
        .AssertEqual actual, expected
    .ShowPendingResult

    .It "should indicate a root-level key doesn't exist"
        actual = suh.KeyExists( suh.HKCU, "Soft ware" )
        expected = False
        .AssertEqual actual, expected
    .ShowPendingResult

    .It "should expand"
        actual = suh.Expand("%UserProfile%")
        expected = sh.ExpandEnvironmentStrings("%UserProfile%")
        .AssertEqual actual, expected

    suh.RegisterWscs

End With
