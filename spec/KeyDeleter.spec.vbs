Option Explicit
Setup
With New TestingFramework
    .describe "KeyDeleter class"
    .it "should instantiate"
        On Error Resume Next
            Set kd = CreateObject("VBScripting.KeyDeleter")
            
            'don't actually delete, for now
            kd.Delete = False
            
            ' "mock" delete shouldn't raise an error
            kd.DeleteKey HKCU, testKeys(0)
            .AssertEqual "{" & Err.Description & "}", "{}"
        On Error Goto 0
    .it "should raise an error on invalid root"
        On Error Resume Next
            kd.DeleteKey -1, testKeys(0)
            .AssertEqual Err.Description, "Expected one of " & validRootsString
        On Error Goto 0
    .it "should allow all valid roots"
        On Error Resume Next
            Dim root : For Each root In validRoots
                kd.ValidateRoot root
            Next
            .AssertEqual "{" & Err.Source & " : " & Err.Description & "}", "{ : }"
        On Error Goto 0
    .it "should ensure BackslashCount(subkey) - BackslashCount(key) = 1"
        On Error Resume Next
            kd.ValidateBackslashCount "v\c\f", "g\r\r"
            .AssertEqual Err.Description, "Expected subkey to have one more backslash than its parent key."
        On Error Goto 0
    .it "should build the subkey correctly"
        kd.DeleteKey HKCU, testKeys(0)
        .AssertEqual kd.SavedKey & "|" & kd.SavedSubkey, "TestKeyDeleter\CLSID\{ADCEC089-0000-11D7-86BF-00606744568C}\subkey|TestKeyDeleter\CLSID\{ADCEC089-0000-11D7-86BF-00606744568C}\subkey\subsubkey"
    .it "should ensure key is not empty"
        On Error Resume Next
            kd.DeleteKey HKCU, ""
            .AssertEqual Err.Description, "The key value is empty, consists of whitespace, or has leading or trailing whitespace."
        On Error Goto 0
    .it "should ensure key is not whitespace"
        On Error Resume Next
            kd.DeleteKey HKCU, " "
            .AssertEqual Err.Description, "The key value is empty, consists of whitespace, or has leading or trailing whitespace."
        On Error Goto 0
    .it "should ensure key is not lead by whitespace"
        On Error Resume Next
            kd.DeleteKey HKCU, " key name"
            .AssertEqual Err.Description, "The key value is empty, consists of whitespace, or has leading or trailing whitespace."
        On Error Goto 0
    .it "should ensure key is not trailed by whitespace"
        On Error Resume Next
            kd.DeleteKey HKCU, "key name "
            .AssertEqual Err.Description, "The key value is empty, consists of whitespace, or has leading or trailing whitespace."
        On Error Goto 0
    .it "should check backslash count"
        .AssertEqual kd.MaxCount, 4
    .it "should ensure that Delete is Boolean"
        On Error Resume Next
            kd.Delete = "a string"
            .AssertEqual Err.Description, "Expected a Boolean."
        On Error Goto 0

    'actually delete now
    kd.Delete = True
    
    .it "should delete a key with subkeys"
        On Error Resume Next
            kd.DeleteKey HKCU, testKeys(0)
            Dim x : x = sh.RegRead("HKCU\" & testKeys(0) & "\")
            .AssertEqual Err.Description, "Invalid root in registry key ""HKCU\TestKeyDeleter\CLSID\{ADCEC089-0000-11D7-86BF-00606744568C}\""."
        On Error Goto 0
    .it "should save the expected success code"
        .AssertEqual kd.Result, 0
    .it "should save the expected code on attempt to delete non-existent key"
        kd.DeleteKey HKCU, testKeys(0)
        .AssertEqual kd.Result, 2
End With

TearDown

Private HKCR, HKCU, HKLM, HKU, HKCC

Private kd
Private sh
Private testKeys

Private validRoots
Private validRootsString
Private validRootsStrings

Sub Setup
    HKCR = &H80000000
    HKCU = &H80000001
    HKLM = &H80000002
    HKU = &H80000003
    HKCC = &H80000005
    'create the test keys
    testKeys = Array( _
        "TestKeyDeleter\CLSID\{ADCEC089-0000-11D7-86BF-00606744568C}", _
        "TestKeyDeleter\Wow6432Node\CLSID\{ADCEC089-0000-11D7-86BF-00606744568C}", _
        "TestKeyDeleter\includer_0000" _
    )
    Set sh = CreateObject("WScript.Shell")
    Dim i : For i = 0 To UBound(testKeys)
        sh.RegWrite "HKCU\" & testKeys(i) & "\subkey\subsubkey\", "test"
    Next
    With CreateObject("Includer")
        ExecuteGlobal .Read("TestingFramework")
    End With
    validRoots = Array( _
       HKCR, HKCU, HKLM, HKU, HKCC)
    validRootsStrings = Array( _
       "-2147483648 = &H80000000 (HKCR)", _
       "-2147483647 = &H80000001 (HKCU)", _
       "-2147483646 = &H80000002 (HKLM)", _
       "-2147483645 = &H80000003 (HKU)", _
       "-2147483643 = &H80000005 (HKCC)")
    validRootsString = vbLf & _
        Join(validRootsStrings, vbLf) & vbLf
End Sub

Sub TearDown
    kd.DeleteKey HKCU, "TestKeyDeleter"
    Set kd = Nothing
    Set sh = Nothing
End Sub