'KeyDeleter class integration test

Option Explicit
Private kd 'KeyDeleter object: what is being tested.
Private incl 'VBScripting.Includer object
Private sh 'WScript.Shell object
Private testKeys 'array of strings: registry key paths
Private HKCR, HKCU, HKLM, HKU, HKCC 'integers: each represents the root of a registry hive
Private validRoots 'array of integers
Private validRootsString 'string of array elements to match an Err.Description
Private validRootsStrings 'array of strings descriptive of registry root integers
Private i 'integer: loop iterator
Private root 'integer: item of items in a For Each loop
Private x 'holds value that would be returned by the RegRead method if the method attempted to read a valid registry location

HKCR = &H80000000
HKCU = &H80000001
HKLM = &H80000002
HKU = &H80000003
HKCC = &H80000005
validRoots = Array( _
    HKCR, HKCU, HKLM, HKU, HKCC )
validRootsStrings = Array( _
    "-2147483648 = &H80000000 (HKCR)", _
    "-2147483647 = &H80000001 (HKCU)", _
    "-2147483646 = &H80000002 (HKLM)", _
    "-2147483645 = &H80000003 (HKU)", _
    "-2147483643 = &H80000005 (HKCC)" )
validRootsString = vbLf & Join(validRootsStrings, vbLf) & vbLf

'Create the test keys.
Set sh = CreateObject( "WScript.Shell" )
testKeys = Array( _
    "TestKeyDeleter\CLSID\{ADCEC089-0000-11D7-86BF-00606744568C}", _
    "TestKeyDeleter\Wow6432Node\CLSID\{ADCEC089-0000-11D7-86BF-00606744568C}", _
    "TestKeyDeleter\includer_0000" _
)
For i = 0 To UBound(testKeys)
    sh.RegWrite "HKCU\" & testKeys(i) & "\subkey\subsubkey\", "test"
Next

Set incl = CreateObject( "VBScripting.Includer" )
Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "KeyDeleter class"
        Set kd = CreateObject( "VBScripting.KeyDeleter" )

    'don't actually delete, for now
    kd.Delete = False

    .it "should not raise an error on mock delete"
        On Error Resume Next
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
            For Each root In validRoots
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
        kd.DeleteKey HKCU, testKeys(0)
        On Error Resume Next
            'Attempt to read from the key just deleted, then assert that the Err.Description of the resulting error is consistent with an attempt to read from a non-existent key.
            x = sh.RegRead( "HKCU\" & testKeys(0) & "\" )
            .AssertEqual Err.Description, "Invalid root in registry key ""HKCU\TestKeyDeleter\CLSID\{ADCEC089-0000-11D7-86BF-00606744568C}\""."
        On Error Goto 0

    .it "should save the expected success code"
        .AssertEqual kd.Result, 0

    .it "should save the expected code on attempt to delete non-existent key"
        kd.DeleteKey HKCU, testKeys(0)
        .AssertEqual kd.Result, 2
End With

kd.DeleteKey HKCU, "TestKeyDeleter"
Set kd = Nothing
Set sh = Nothing
