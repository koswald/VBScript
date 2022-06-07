Set sh = CreateObject( "WScript.Shell" )
Set fso = CreateObject( "Scripting.FileSystemObject" )
With CreateObject( "VBScripting.Includer" )
    Execute .Read( "VBSLogger" )
End With
Set logger = New VBSLogger
response = sh.PopUp("test", 2, "PopUp bug", vbOKCancel)
Set stream = fso.OpenTextFile("fixture\PopUp2.fixture.txt", ForWriting, CreateNew)
stream.WriteLine response

Set fso = Nothing
Set sh = Nothing
Set logger = Nothing
Const CreateNew = True
Const ForWriting = 2
