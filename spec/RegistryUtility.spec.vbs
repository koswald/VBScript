
'test RegistryUtility.vbs

With CreateObject("includer")
    Execute(.read("RegistryUtility"))
    Execute(.read("TestingFramework"))
    Execute(.read("VBSNatives"))
End With

Dim r : Set r = New RegistryUtility 'class under test
Dim n : Set n = New VBSNatives

Dim subKey : subKey = "Software\Scripts by Karl"
Dim valueName : valueName = "" 'use the subKey's default value
Dim value : value = n.fso.GetTempName 'a random string
Dim key : key = "HKCU\" & subKey & "\" & valueName 'registry key format used by WScript.Shell.RegRead & .RegWrite; this format is not used by the class under test

'create the test registry key, if it doesn't exist;
'if it does exist, then save the existing value

Dim savedValue : savedValue = ""
On Error Resume Next
    savedValue = n.sh.RegRead(key)
    If Err Then n.sh.RegWrite key, ""
On Error Goto 0

With New TestingFramework

    .describe "RegistryUtility class"

    .it "should write a registry string (REG_SZ) value"

        r.SetStringValue r.HKCU, subKey, valueName, value

        .AssertEqual n.sh.RegRead(key), value

    .it "should read a registry string (REG_SZ) value"

        .AssertEqual r.GetStringValue(r.HKCU, subKey, valueName), value

    .it "should access a registry by computer name"

        With CreateObject("WScript.Network")
            r.SetPC .ComputerName
        End With

        .AssertEqual r.GetStringValue(r.HKCU, subKey, valueName), value

End With

'restore the saved registry value

n.sh.RegWrite key, savedValue
