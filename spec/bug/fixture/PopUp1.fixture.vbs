Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set incl = CreateObject("VBScripting.Includer")
Execute incl.Read("VBSLogger")
Set logger = New VBSLogger
response = sh.PopUp("test", 2, "PopUp bug", vbOKCancel)
Set stream = fso.OpenTextFile("fixture\PopUp1.fixture.txt", ForWriting, CreateNew)
stream.WriteLine response

Set fso = Nothing
Set sh = Nothing
Set logger = Nothing
Const CreateNew = True
Const ForWriting = 2
