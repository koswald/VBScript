
'test RegistryUtility.vbs with elevated privileges

With CreateObject("includer")
    Execute(.read("RegistryUtility"))
    Execute(.read("TestingFramework"))
End With

With New TestingFramework

    .describe "RegistryUtility class, elevated privileges"

        Dim r : Set r = New RegistryUtility

    .it "should read a registry expanded string (REG_EXPAND_SZ) value"

        subKey = "WSFFile\Shell\Open\Command"

        .AssertEqual CBool(InStr(r.GetExpandedStringValue(r.HKCR, subKey, ""), "Script.exe""")), True

    .it "should return a string showing a reg value type REG_EXPAND_SZ"

        .AssertEqual r.GetRegValueTypeString(r.HKCR, subKey, ""), "REG_EXPAND_SZ"
End With
