
'test WoWChecker.vbs

With CreateObject("includer")
    Execute(.read("TestingFramework"))
    Execute(.read("WoWChecker"))
    Execute(.read("VBSNatives"))
End With

Dim n : Set n = New VBSNatives

With New TestingFramework

    .describe "WoWChecker class"

        Dim chkr : Set chkr = New WoWChecker

    .it "should return False, with a correctly configured test"

        .AssertEqual chkr.isWoW, False

    .it "should return an obj self reference on ByCheckSum call"

        Set chkr = New WowChecker.ByCheckSum

        .AssertEqual chkr.isSysWoW64, False

    .it "should return an obj self reference on BySize call"

        Set chkr = New WoWChecker.BySize

        .AssertEqual chkr.isSystem32, True

    .it "should have a public default property"

        .AssertEqual chkr, False

    .it "should return True with a 32-bit process on isWoW call"

        Dim pipe : Set pipe = n.sh.Exec("%SystemRoot%\SysWoW64\cscript.exe //nologo fixture\WoWChecker.GetWoW.vbs")

        .AssertEqual CBool(pipe.StdOut.ReadLine), True

    .it "should return False with a 64-bit process on isWoW call"

        Set pipe = n.sh.Exec("%SystemRoot%\System32\cscript.exe //nologo fixture\WoWChecker.GetWoW.vbs")

        .AssertEqual CBool(pipe.StdOut.ReadLine), False
End With

'garbage collection

Set chkr = Nothing
Set pipe = Nothing
