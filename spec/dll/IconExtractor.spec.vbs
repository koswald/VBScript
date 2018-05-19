Option Explicit : Initialize
With New TestingFramework

    .describe "IconExtractor class"
        Dim extractor
        Set extractor = CreateObject("VBScripting.IconExtractor")

    .it "should get the number of icons in a file"
        .AssertEqual extractor.IconCount("%SystemRoot%\System32\imageres.dll"), 412

    .it "should extract an icon and save it"
        Dim resFile : resFile = "%SystemRoot%\System32\imageres.dll"
        Dim icoFile : icoFile = "%UserProfile%\Desktop\test.ico"
        extractor.Save resFile, 0, icoFile, True
        .AssertEqual fso.GetFile(Expand(icoFile)).Size > 0, True

    .DeleteFile icoFile
    Teardown
End With

Dim fso, sh, includer
Sub Initialize
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sh = CreateObject("WScript.Shell")
    Set includer = CreateObject("VBScripting.Includer")
    ExecuteGlobal includer.Read("TestingFramework")
End Sub
Sub Teardown
    Set fso = Nothing
    Set sh = Nothing
    Set includer = Nothing
End Sub
Function Expand(str)
    Expand = sh.ExpandEnvironmentStrings(str)
End Function
