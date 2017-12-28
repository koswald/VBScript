
'test PrivilegeChecker with elevated privileges

Option Explicit : Initialize

With CreateObject("includer")
    Execute .read("TestingFramework")
    Execute .read("PrivilegeChecker")
End With

With New TestingFramework

    .describe "PrivilegeChecker class"

    .it "should indicate that privileges are elevated"
        .AssertEqual pc, True

End With

Dim pc

Sub Initialize
    With CreateObject("includer")
        Execute .read("PrivilegeChecker")
        Execute .read("VBSApp")
    End With
    Set pc = New PrivilegeChecker
    Dim app : Set app = New VBSApp
    If Not pc Then
        app.SetUserInteractive False
        app.RestartWith "cscript.exe", "/k", True
    End If
End Sub