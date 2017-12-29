
'test includer.wsc

'GetObj method under development

Option Explicit : Initialize

With New TestingFramework

    .describe "includer.wsc dependency manager scriptlet"
        Dim incl : Set incl = CreateObject("includer")

    Dim i : For i = 1 To UBound(passing)
        .it format(Array(_
            "should return an object on .GetObj(""%s"")", passing(i)))
        .AssertEqual TypeName(incl.GetObj(passing(i))), passing(i)
    Next

    Dim obj
    For i = 1 To UBound(failing)
        .it format(Array(_
            "should not raise an error on .GetObj(""%s"")", failing(i)))
            On Error Resume Next
                Set obj = incl.GetObj(failing(i))
                .AssertEqual format(Array( _
                    "{%s}", Err.Description _
                )), "{}"
            On Error Goto 0
    Next
End With

'cleanup
Set fso = Nothing

Dim fso
Dim format
Dim passing, failing

Sub Initialize
    With CreateObject("includer")
        Execute .read("StringFormatter")
        ExecuteGlobal .read("TestingFramework")
    End With
    Set format = New StringFormatter
    Set fso = CreateObject("Scripting.FileSystemObject")
    passing = Array("" _
        , "CommandParser" _
        , "EncodingAnalyzer" _
        , "Function123" _
        , "GUIDGenerator" _
        , "MathConstants" _
        , "PrivilegeChecker" _
        , "RegExFunctions" _
        , "RegistryUtility" _
        , "ShellConstants" _
        , "SpecialFolders" _
        , "StreamConstants" _
        , "StringFormatter" _
        , "TextStreamer" _
        , "TimeFunctions" _
        , "VBSArguments" _
        , "VBSArrays" _
        , "VBSClipboard" _
        , "VBSEnvironment" _
        , "VBSEventLogger" _
        , "VBSExtracter" _
        , "VBSFileSystem" _
        , "VBSHoster" _
        , "VBSLogger" _
        , "VBSMessages" _
        , "VBSPower" _
        , "VBSTestRunner" _
        , "VBSTimer" _
        , "VBSTroubleshooter" _
        , "VBSValidator" _
        , "WMIUtility" _
        , "WoWChecker" _
    )
    failing = Array("" _
        , "Chooser" _
        , "DotNetCompiler" _
        , "HTAApp" _
        , "TestingFramework" _
        , "VBSApp" _
        , "WindowsUpdatesPauser" _
    )
End Sub