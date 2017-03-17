'test DotNetCompiler class

With CreateObject("includer")
    Execute(.read("DotNetCompiler"))
    Execute(.read("TestingFramework"))
End With

With New TestingFramework

    .describe "DotNetCompiler class"
        Dim dnc : Set dnc = New DotNetCompiler

    'setup
        Const verifyKeyDeletion = False
        Const synchronous = True
        Const hidden = 0
        Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim sh : Set sh = CreateObject("WScript.Shell")

        Dim sourceFile1 : sourceFile1 = "fixture\.net\SourceCode1.cs"
        Dim baseName1 : baseName1 = fso.GetBaseName(sourceFile1)
        Dim sourceFolder1 : sourceFolder1 = fso.GetParentFolderName(fso.GetAbsolutePathName(sourceFile1))
        Dim sourceBase1 : sourceBase1 = sourceFolder1 & "\" & baseName1

        Dim sourceFile2 : sourceFile2 = "fixture\.net\SourceCode2.cs"
        Dim baseName2 : baseName2 = fso.GetBaseName(sourceFile2)
        Dim sourceFolder2 : sourceFolder2 = fso.GetParentFolderName(fso.GetAbsolutePathName(sourceFile2))
        Dim sourceBase2 : sourceBase2 = sourceFolder2 & "\" & baseName2

        Dim keysToRemove : keysToRemove = Array( _
            "HKCR\Wow6432Node\CLSID\{2650C2AD-1AF8-495F-AB4D-6C61BD463EA4}", _
            "HKCR\Wow6432Node\CLSID\{2650C2AD-1BF8-495F-AB4D-6C61BD463EA4}", _
            "HKCR\Wow6432Node\CLSID\{2650C2AD-2AF8-495F-AB4D-6C61BD463EA4}", _
            "HKCR\Wow6432Node\CLSID\{2650C2AD-2BF8-495F-AB4D-6C61BD463EA4}", _
            "HKCR\CLSID\{2650C2AD-1AF8-495F-AB4D-6C61BD463EA4}", _
            "HKCR\CLSID\{2650C2AD-1BF8-495F-AB4D-6C61BD463EA4}", _
            "HKCR\CLSID\{2650C2AD-2AF8-495F-AB4D-6C61BD463EA4}", _
            "HKCR\CLSID\{2650C2AD-2BF8-495F-AB4D-6C61BD463EA4}")
        Dim key1a_x86 : key1a_x86 = 0 'progid    for 32-bit sourceFile1
        Dim key1b_x86 : key1b_x86 = 1 'interface for 32-bit sourceFile1
        Dim key2a_x86 : key2a_x86 = 2 'progid    for 32-bit sourceFile2
        Dim key2b_x86 : key2b_x86 = 3 'interface for 32-bit sourceFile2
        Dim key1a_x64 : key1a_x64 = 4 'progid    for 64-bit sourceFile1
        Dim key1b_x64 : key1b_x64 = 5 'interface for 64-bit sourceFile1
        Dim key2a_x64 : key2a_x64 = 6 'progid    for 64-bit sourceFile2
        Dim key2b_x64 : key2b_x64 = 7 'interface for 64-bit sourceFile2

        Dim filesToRemove : filesToRemove = Array( _
            sourceBase1 & ".snk", _
            sourceBase1 & ".dll", _
            sourceBase2 & ".snk", _
            sourceBase2 & ".dll")
        Dim iSnkFile1 : iSnkFile1 = 0
        Dim iDllFile1_x64 : iDllFile1_x64 = 1
        Dim iSnkFile2 : iSnkFile2 = 2
        Dim iDllFile2_x64 : iDllFile2_x64 = 3

        Dim foldersToRemove : foldersToRemove = Array( _
            sourceFolder1 & "\createTargetFolderTest", _
            sourceFolder1 & "\lib\64", _
            sourceFolder1 & "\lib\32")
        Dim iCreateFolderTestFolder1 : iCreateFolderTestFolder1 = 0
        Dim iTargetFolder1_x64 : iTargetFolder1_x64 = 1
        Dim iTargetFolder1_x86 : iTargetFolder1_x86 = 2

        Cleanup 'remove junk from past erring tests, if any
        EnsureNoJunk
        dnc.SetUserInteractive False
        dnc.OnUserCancelQuitScript = True

    .it "should fail to create a key pair without a sourceFile or targetName"
        On Error Resume Next
            dnc.GenerateKeyPair
            .AssertErrorRaised
        On Error Goto 0

    .it "should create a strong name key pair"
        dnc.SetSourceFile sourceFile1
        dnc.GenerateKeyPair

        .AssertEqual fso.FileExists(filesToRemove(iSnkFile1)), True

    .it "should set the target folder"
        dnc.SetTargetFolder foldersToRemove(iCreateFolderTestFolder1)
        .AssertEqual Err, 0

    .it "should create the target folder"
        .AssertEqual fso.FolderExists(foldersToRemove(iCreateFolderTestFolder1)), True

    .it "should compile a .cs file"
        dnc.Compile
        .AssertEqual fso.FileExists(filesToRemove(iDllFile1_x64)), True

    .it "should register a .dll in a custom location"
        dnc.SetTargetFolder foldersToRemove(iTargetFolder1_x64)
        dnc.Register
        On Error Resume Next
            EnsureKeyExists keysToRemove(key1a_x64)
            .AssertEqual Err & ": " & Err.Description, "0: "
        On Error Goto 0

    .it "should compile and register a 32-bit .dll"
        dnc.SetTargetFolder foldersToRemove(iTargetFolder1_x86)
        dnc.SetTargetName baseName1 & "32"
        dnc.SetBitness 32
        dnc.Compile
        dnc.Register
        On Error Resume Next
            EnsureKeyExists keysToRemove(key1a_x86)
            .AssertEqual Err & ": " & Err.Description, "0: "
        On Error Goto 0

    .it "should add a reference"
    .it "should unregister"
End With

'teardown
    Cleanup
    EnsureNoJunk
    Quit

Sub Quit
    CollectGarbage
    WScript.Quit
End Sub

'delete selected files, folders, and keys
Sub Cleanup
    Dim i
    For i = 0 To UBound(filesToRemove)
        If fso.FileExists(filesToRemove(i)) Then fso.DeleteFile filesToRemove(i), True
    Next
    For i = 0 To UBound(foldersToRemove)
        If fso.FolderExists(foldersToRemove(i)) Then fso.DeleteFolder foldersToRemove(i), True
    Next
    For i = 0 To UBound(keysToRemove)
        If verifyKeyDeletion Then If vbCancel = MsgBox("Delete " & keysToRemove(i) & "?", vbOKCancel + vbQuestion, WScript.ScriptName) Then Exit For
        If keyExists(keysToRemove(i)) Then sh.Run "reg delete " & keysToRemove(i) & " /f", hidden, synchronous
    Next
End Sub

'Raise an error if any of the specified files, folders, or keys
'were not removed
Sub EnsureNoJunk
    Dim i
    For i = 0 To UBound(filesToRemove)
        If fso.FileExists(filesToRemove(i)) Then Err.Raise 1, WScript.ScriptName, "Could not clean up file " & fso.GetAbsolutePathName(filesToRemove(i))
    Next
    For i = 0 To UBound(foldersToRemove)
        If fso.FolderExists(foldersToRemove(i)) Then Err.Raise 2, WScript.ScriptName, "Could not clean up folder " & fso.GetAbsolutePathName(foldersToRemove(i))
    Next
    For i = 0 To UBound(keysToRemove)
        If keyExists(keysToRemove(i)) Then Err.Raise 3, WScript.ScriptName, "Could not clean up registry key " & keysToRemove(i)
    Next
End Sub

'Return True if the given registry key exists
Function keyExists(key)
    On Error Resume Next
        EnsureKeyExists key
        keyExists = Not Cbool(Err.Number)
    On Error Goto 0
End Function

'Raise an error if the registry key doesn't exist
Sub EnsureKeyExists(key)
    sh.RegRead key & "\"
End Sub

'release memory associated with selected objects
Sub CollectGarbage
    Set fso = Nothing
    Set sh = Nothing
End Sub
