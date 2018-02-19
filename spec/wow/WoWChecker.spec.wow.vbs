
'test WoWChecker.vbs
'intended to be run with 32-bit cscript.exe

With CreateObject("VBScripting.Includer")
    Execute .read("TestingFramework")
    Execute .read("WoWChecker")
End With

With New TestingFramework

    .describe "WoWChecker class"
        Dim chkr : Set chkr = New WoWChecker

    'setup
        Dim sh : Set sh = CreateObject("WScript.Shell")
        If Not chkr.isWoW Then WScript.StdOut.WriteLine "This test must be launched with the 32-bit cscript.exe." : WScript.Quit

    .it "should return an obj self reference on ByCheckSum call"
        Set chkr = New WowChecker.ByCheckSum
        .AssertEqual chkr.isSysWoW64, True

    .it "should return an obj self reference on BySize call"
        Set chkr = New WoWChecker.BySize
        .AssertEqual chkr.isSystem32, False

    .it "should have a default property"
        .AssertEqual chkr, True

    .it "should return True with a 32-bit process on isWoW call"
        Dim pipe : Set pipe = sh.Exec("%SystemRoot%\SysWoW64\cscript.exe //nologo fixture\WoWChecker.GetWoW.vbs")
        .AssertEqual pipe.StdOut.ReadLine, "True"

    .it "should return True with a 64-bit process on isWoW call"
        Set pipe = sh.Exec("%SystemRoot%\System32\cscript.exe //nologo fixture\WoWChecker.GetWoW.vbs")
        .AssertEqual pipe.StdOut.ReadLine, "True"

End With

'garbage collection
Set chkr = Nothing
Set pipe = Nothing
Set sh = Nothing
