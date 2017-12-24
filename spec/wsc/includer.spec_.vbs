
'GetObj method under development

Dim incl : Set incl = CreateObject("includer")
Dim tester : Set tester = incl.GetObj("TestingFramework")

MsgBox "VarType: " & VarType(tester) & vbLf & "TypeName: " & TypeName(tester)

With tester

    .describe "includer.wsc dependency manager scriptlet"

    .it "should return an object instance given the class name"
        .AssertEqual True, True 'GetObj method returned an instance of TestingFramework
End With
