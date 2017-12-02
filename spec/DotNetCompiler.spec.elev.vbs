'test DotNetCompiler class
'intended to be run with elevated privileges

With CreateObject("includer")
    Execute .read("DotNetCompiler")
    Execute .read("TestingFramework")
    Execute .read("StringFormatter")
    Execute .read("VBSLogger")
End With

With New TestingFramework

    .describe "DotNetCompiler class"
        Dim dnc : Set dnc = New DotNetCompiler

    'setup
        Const verifyKeyDeletion = False
        Const synchronous = True
        Const hidden = 0

        dnc.SetUserInteractive False 'set to True for debugging

        Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim sh : Set sh = CreateObject("WScript.Shell")
        Dim format : Set format = New StringFormatter
        Dim log : Set log = New VBSLogger

        Dim sourceFile1 : sourceFile1 = "fixture\.net\SourceCode1.cs"
        Dim baseName1 : baseName1 = fso.GetBaseName(sourceFile1)
        Dim sourceFolder1 : sourceFolder1 = fso.GetParentFolderName(fso.GetAbsolutePathName(sourceFile1))
        Dim sourceBase1 : sourceBase1 = sourceFolder1 & "\" & baseName1

        Dim sourceFile2 : sourceFile2 = "fixture\.net\SourceCode2.cs"
        Dim baseName2 : baseName2 = fso.GetBaseName(sourceFile2)
        Dim sourceFolder2 : sourceFolder2 = fso.GetParentFolderName(fso.GetAbsolutePathName(sourceFile2))
        Dim sourceBase2 : sourceBase2 = sourceFolder2 & "\" & baseName2

        Dim testKeys : testKeys = Array( _
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

        Dim testFiles : testFiles = Array( _
            "fixture\.net\DotNetCompiler.snk", _
            sourceBase1 & ".dll", _
            "fixture\.net\DotNetCompiler.snk", _
            sourceBase2 & ".dll")
        Dim iSnkFile1 : iSnkFile1 = 0
        Dim iDllFile1_x64 : iDllFile1_x64 = 1
        Dim iSnkFile2 : iSnkFile2 = 2
        Dim iDllFile2_x64 : iDllFile2_x64 = 3

        Dim testFolders : testFolders = Array( _
            sourceFolder1 & "\" & GetPCName & "\createTargetFolderTest", _
            sourceFolder1 & "\" & GetPCName & "\lib\64", _
            sourceFolder1 & "\" & GetPCName & "\lib\32")
        Dim iCreateFolderTestFolder1 : iCreateFolderTestFolder1 = 0
        Dim iTargetFolder1_x64 : iTargetFolder1_x64 = 1
        Dim iTargetFolder1_x86 : iTargetFolder1_x86 = 2

        Cleanup 'remove junk from past erring tests, if any
        EnsureCleanupWorked

    .it "should fail to create a key pair without a sourceFile or targetName"
        On Error Resume Next
            dnc.GenerateKeyPair
            .AssertErrorRaised
        On Error Goto 0

    .it "should create a strong name key pair"
        dnc.SetSourceFile sourceFile1
        dnc.SetKeyFile testFiles(iSnkFile1)
        dnc.GenerateKeyPair
        .AssertEqual fso.FileExists(testFiles(iSnkFile1)), True

    .it "should set the target folder"
        dnc.SetTargetFolder testFolders(iCreateFolderTestFolder1)
        .AssertEqual Err, 0

    .it "should create the target folder"
        .AssertEqual fso.FolderExists(testFolders(iCreateFolderTestFolder1)), True

    .it "should compile a .cs file"
        dnc.Compile
        .AssertEqual fso.FileExists(testFiles(iDllFile1_x64)), True

    .it "should register a .dll in a custom location"
        dnc.SetTargetFolder testFolders(iTargetFolder1_x64)
        dnc.Register
        On Error Resume Next
            EnsureKeyExists(testKeys(key1a_x64))
            .AssertEqual format(Array("%s: %s", Err, Err.Description)), "0: "
        On Error Goto 0

    .it "should compile and register a 32-bit .dll"
        dnc.SetTargetFolder testFolders(iTargetFolder1_x86)
        dnc.SetTargetName baseName1 & "_32"
        dnc.SetBitness 32
        dnc.Compile
        dnc.Register
        On Error Resume Next
            EnsureKeyExists testKeys(key1a_x86)
            .AssertEqual format(Array("%s: %s", Err, Err.Description)), "0: "
        On Error Goto 0

    .it "should add a reference"
        dnc.SetSourceFile sourceFile2
        dnc.GenerateKeyPair
        dnc.SetTargetFolder testFolders(iTargetFolder2_x86)
        dnc.SetTargetName dnc.GetTargetName & "_32"
        dnc.AddRef "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\System.Speech.dll"
        dnc.Compile
        dnc.Register
        On Error Resume Next
            EnsureKeyExists testKeys(key2a_x86)
            .AssertEqual Err & ": " & Err.Description, "0: "
        On Error Goto 0

    .it "should unregister"
        dnc.Unregister
        On Error Resume Next
            EnsureKeyIsGone(testKeys(key2a_x86))
            EnsureKeyIsGone(testKeys(key2b_x86))
            .AssertEqual Err & ": " & Err.Description, "0: "
        On Error Goto 0

End With

'teardown
    Cleanup
    EnsureCleanupWorked
    Quit

Sub Quit
    EmptyTheTrash
    WScript.Quit
End Sub

'delete selected files, folders, and keys
Sub Cleanup
    Dim i
    For i = 0 To UBound(testFiles)
        If fso.FileExists(testFiles(i)) Then fso.DeleteFile testFiles(i), True
    Next
    For i = 0 To UBound(testFolders)
        RemoveFolder(testFolders(i))
    Next
    For i = 0 To UBound(testKeys)
        If verifyKeyDeletion Then If vbCancel = MsgBox("Delete " & testKeys(i) & "?", vbOKCancel + vbQuestion, WScript.ScriptName) Then Quit
        If keyExists(testKeys(i)) Then sh.Run "reg delete " & testKeys(i) & " /f", hidden, synchronous
    Next
End Sub

Sub RemoveFolder(folder)
    If Not fso.FolderExists(folder) Then Exit Sub
    On Error Resume Next
        fso.DeleteFolder folder, True
        If Err Then
            Dim msg : msg = format(Array( _
                "Error %s ( %s ) removing folder %s", _
                Err.Number, Err.Description, folder))
            log msg
        End If
    On Error Goto 0
End Sub

'Raise an error if any of the specified files,
'folders, or keys were not removed
Sub EnsureCleanupWorked
    Dim i
    For i = 0 To UBound(testFiles)
        If fso.FileExists(testFiles(i)) Then Err.Raise 1, WScript.ScriptName, "Could not clean up file " & fso.GetAbsolutePathName(testFiles(i))
    Next
    For i = 0 To UBound(testFolders)
        If fso.FolderExists(testFolders(i)) Then Err.Raise 2, WScript.ScriptName, "Could not clean up folder " & fso.GetAbsolutePathName(testFolders(i))
    Next
    For i = 0 To UBound(testKeys)
        EnsureKeyIsGone(testKeys(i))
    Next
End Sub

Sub EnsureKeyIsGone(key)
    If keyExists(key) Then Err.Raise 3, WScript.ScriptName, "Could not clean up registry key " & key
End Sub

'Return True if the given registry key exists
Function keyExists(key)
    On Error Resume Next
        EnsureKeyExists key
        keyExists = Not Cbool(Err.Number)
    On Error Goto 0
End Function

'Raise an error if the registry key doesn't exist
Function EnsureKeyExists(key)
    EnsureKeyExists = sh.RegRead(key & "\")
End Function

'Get the computer name
Function GetPCName
    With CreateObject("WScript.Network")
        GetPCName = .ComputerName
    End With
End Function

'release memory associated with selected objects
Sub EmptyTheTrash
    Set fso = Nothing
    Set sh = Nothing
End Sub
