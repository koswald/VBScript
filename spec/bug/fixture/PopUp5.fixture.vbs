Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
With CreateObject("VBScripting.Includer")
    Execute .Read("TextStreamer")
End With
Set streamer = New TextStreamer
response = sh.PopUp("test", 2, "PopUp bug", vbOKCancel)
Set stream = fso.OpenTextFile("fixture\PopUp5.fixture.txt", ForWriting, CreateNew)
stream.WriteLine response

Set fso = Nothing
Set sh = Nothing
Set streamer = Nothing
Const CreateNew = True
Const ForWriting = 2
