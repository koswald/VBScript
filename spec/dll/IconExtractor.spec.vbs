Option Explicit : Initialize
With New TestingFramework

    .describe "IconExtractor class"
        Dim extractor
        Set extractor = CreateObject("VBScripting.IconExtractor")

    .it "should get the number of icons in " & resFile
        .AssertEqual extractor.IconCount(resFile), 334

    .it "should extract an icon and save it"
        extractor.Save resFile, 289, icoFile, True
        .AssertEqual fso.GetFile(Expand(icoFile)).Size > 0, True

    .DeleteFile icoFile
End With

Teardown

Function Expand(str)
    Expand = sh.ExpandEnvironmentStrings(str)
End Function

Dim resFile, icoFile
Dim fso, sh, includer
Sub Initialize
    resFile = "%SystemRoot%\System32\imageres.dll"
    icoFile = "%UserProfile%\Desktop\IconExtractor.spec.ico"
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
