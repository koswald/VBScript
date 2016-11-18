
'test RegistryUtility.vbs

With CreateObject("includer")
    Execute(.read("RegistryUtility"))
    Execute(.read("TestingFramework"))
    Execute(.read("VBSNatives"))
End With

With New TestingFramework

    .describe "RegistryUtility class"

        Dim r : Set r = New RegistryUtility

    'setup

        Dim n : Set n = New VBSNatives
        Dim subKey : subKey = "Software\VBScripts"
        Dim valueName : valueName = "" 'use the subKey's default value
        Dim valueName2 : valueName2 = n.fso.GetTempName
        Dim value : value = n.fso.GetTempName
        Dim value2 : value2 = n.fso.GetTempName
        Dim key : key = "HKCU\" & subKey & "\" & valueName 'registry key format used by WScript.Shell.RegRead & .RegWrite; this format is not used by the class under test
        Dim key2 : key2 = "HKCU\" & subKey & "\" & valueName2

        'create the test registry keys, if they doesn't exist;
        'save the existing value for the first key

        Dim savedValue : savedValue = ""
        On Error Resume Next
            savedValue = n.sh.RegRead(key)
            If Err Then n.sh.RegWrite key, ""
            n.sh.RegWrite key2, value2, "REG_EXPAND_SZ"
        On Error Goto 0

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

    .it "should return an integer showing a reg value type"

        .AssertEqual r.GetRegValueType(r.HKCU, subKey, valueName), r.REG_SZ

    .it "should return a string showing a reg value type REG_SZ"

        .AssertEqual r.GetRegValueTypeString(r.HKCU, subKey, valueName), "REG_SZ"

    .it "should read a registry expanded string (REG_EXPAND_SZ) value"

        .AssertEqual CBool(InStr(r.GetStringValue(r.HKCU, subKey, valueName2), value2)), True

    .it "should return a string showing a reg value type REG_EXPAND_SZ"

        .AssertEqual r.GetRegValueTypeString(r.HKCU, subKey, valueName2), "REG_EXPAND_SZ"
End With

'restore the saved registry value

n.sh.RegWrite key, savedValue

'delete the second key

n.sh.RegDelete key2
