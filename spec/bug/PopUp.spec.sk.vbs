'reproduce the VBScripting bug with WshShell.PopUp
Set includer = CreateObject("VBScripting.Includer")
Execute includer.Read("TestingFramework")
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
files = Array(""_
    , "fixture\PopUp1.fixture.vbs" _
    , "fixture\PopUp1.fixture.txt" _
    , "fixture\PopUp2.fixture.vbs" _
    , "fixture\PopUp2.fixture.txt" _
    , "fixture\PopUp3.fixture.wsf" _
    , "fixture\PopUp3.fixture.txt" _
    , "fixture\PopUp4.fixture.wsf" _
    , "fixture\PopUp4.fixture.txt" _
    , "fixture\PopUp5.fixture.vbs" _
    , "fixture\PopUp5.fixture.txt" _
)
With New TestingFramework
    .describe "VBScripting bug with WshShell.PopUp"
    .it "fails to exhibit when preventive syntax is used - .vbs"
        sh.Run files(1)
        If Not .MessageAppeared("PopUp bug", 2, "{Esc}") Then Err.Raise 1,, WScript.ScriptName & ": PopUp1 failed to appear."
        WScript.Sleep pause 'wait for fixture script to write to output file
        .AssertEqual fso.OpenTextFile(files(2)).ReadLine, "2" '2 = vbCancel = response of PopUp dialog to the Esc key
    .it "exhibits when preventive syntax is not used - .vbs"
        sh.Run files(3)
        If Not .MessageAppeared("PopUp bug", 2, "{Esc}") Then Err.Raise 2,, WScript.ScriptName & ": PopUp2 failed to appear."
        WScript.Sleep pause
        .AssertEqual fso.OpenTextFile(files(4)).ReadLine, "-1" '-1 = response of timed out PopUp dialog or, notably, response of the PopUp dialog to the Esc key when preventive syntax is not used
    .it "fails to exhibit when preventive syntax is used - .wsf"
        sh.Run files(5)
        If Not .MessageAppeared("PopUp bug", 2, "{Esc}") Then Err.Raise 3,, WScript.ScriptName & ": PopUp3 failed to appear."
        WScript.Sleep pause
        .AssertEqual fso.OpenTextFile(files(6)).ReadLine, "2"
    .it "fails to exhibit when preventive syntax is not used - .wsf"
        sh.Run files(7)
        If Not .MessageAppeared("PopUp bug", 2, "{Esc}") Then Err.Raise 4,, WScript.ScriptName & ": PopUp4 failed to appear."
        WScript.Sleep pause
        .AssertEqual fso.OpenTextFile(files(8)).ReadLine, "2"
    .it "exhibits when preventive syntax is not used - .vbs - TextStreamer"
        sh.Run files(9)
        If Not .MessageAppeared("PopUp bug", 2, "{Esc}") Then Err.Raise 5,, WScript.ScriptName & ": PopUp5 failed to appear."
        WScript.Sleep pause
        .AssertEqual fso.OpenTextFile(files(10)).ReadLine, "-1"

    .DeleteFiles Array(files(2), files(4), files(6), files(8), files(10))
End With

Const pause = 100
Set includer = Nothing
Set sh = Nothing
Set fso = Nothing

