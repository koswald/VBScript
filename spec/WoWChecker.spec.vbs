'Integration test for the WoWChecker class
'intended to be run with 64-bit cscript.exe

Option Explicit
Dim chkr 'WoWChecker object in test
Dim incl 'VBScripting.Includer object
Dim sh 'WScript.Shell object
Dim errMsg32bitOs 'error message
Dim actual, expected
Dim pipe 'console process

Set incl = CreateObject( "VBScripting.Includer" )
Set sh = CreateObject( "WScript.Shell" )
errMsg32bitOs = "OS is 32-bit. This spec is not applicable."

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "WoWChecker class"
        Execute incl.Read( "WoWChecker" )
        Set chkr = New WoWChecker

    If Not chkr.OSIs64Bit Then
        WScript.StdOut.WriteLine errMsg32bitOs
        WScript.Quit
    End If

    .it "should return False, with a correctly configured test"
        .AssertEqual chkr.isWoW, False

    .it "should return an object self reference on ByCheckSum call"
        Set chkr = New WowChecker.ByCheckSum
        actual = TypeName( chkr )
        expected = "WoWChecker"
        .AssertEqual actual, expected

    .it "should return an object self reference on BySize call"
        Set chkr = New WoWChecker.BySize
        actual = TypeName( chkr )
        expected = "WoWChecker"
        .AssertEqual actual, expected

    .it "should have a default property"
        .AssertEqual chkr, False

    .it "should return True with a 32-bit process on isWoW call"
        Set pipe = sh.Exec("%SystemRoot%\SysWoW64\cscript.exe //nologo fixture\WoWChecker.GetWoW.vbs")
        .AssertEqual pipe.StdOut.ReadLine, "True"

    .it "should return False with a 64-bit process on isWoW call"
        Set pipe = sh.Exec("%SystemRoot%\System32\cscript.exe //nologo fixture\WoWChecker.GetWoW.vbs")
        .AssertEqual pipe.StdOut.ReadLine, "False"

End With

'garbage collection
Set chkr = Nothing
Set pipe = Nothing
Set sh = Nothing
