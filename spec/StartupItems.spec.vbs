'Integration test for the StartupItems class

'Integration test for the StartupItems class.
'Also tests the ArrayOfObjects and NameValue classes.

Option Explicit
Dim si 'StartupItems object: what is to be tested
Dim incl 'VBScripting.Includer object
Dim initialUBound 'integer
Dim actual, expected 'assertion arguments

Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "StartupItems class"
        Execute incl.Read( "StartupItems" )
        Set si = New StartupItems

    .It "should create an item"
        initialUBound = UBound(si.Items)
        'testing the UpdateItem method should also test the CreateItem method
        si.UpdateItem "SISpec1_Test", "notepad """ & WScript.ScriptFullName & """"
        'testing the Items property should also test the Item property
        actual = UBound(si.Items)
        expected = initialUBound + 1
    .AssertEqual actual, expected

    .It "should delete an item"
        'testing the RemoveItem method should also test the DeleteItem method
        si.RemoveItem "SISpec1_Test"
    .AssertEqual UBound(si.Items), initialUBound

    .It "should get the default root for StdRegProv methods"
    .AssertEqual si.Root, &H80000001

    .It "should get the default root for WScript.Shell methods"
    .AssertEqual si.WSHRoot, "HKCU"

    .It "should get the standard branch"
        expected = "Software\Microsoft\Windows\CurrentVersion\Run"
    .AssertEqual si.StandardBranch, expected

    .It "should get the WOW branch"
        expected = "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run"
    .AssertEqual si.WoWBranch, expected

    .It "should set the all-users root for StdRegProv methods"
        si.Root = si.HKLM
    .AssertEqual si.Root, &H80000002

    .It "should get the all-users root for WScript.Shell methods"
    .AssertEqual si.WSHRoot, "HKLM"

    .It "should not (directly) set the all-users root for WScript.Shell methods"
        On Error Resume Next
            si.WSHRoot = "HKLM"
            actual = Left( Err.Description, 45 )
            expected = "Object doesn't support this property or metho"
        On Error Goto 0
    .AssertEqual actual, expected

End With
