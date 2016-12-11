
'test VBSClipboard.vbs

With CreateObject("includer")
	Execute(.read("VBSClipboard"))
	Execute(.read("TestingFramework"))
End With

With New TestingFramework

	.describe "VBSClipboard class"
		Dim cb : Set cb = New VBSClipboard

	'setup
		Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
		
	.it "should copy text to the clipboard"
		Dim randomText : randomText = fso.GetTempName
		cb.SetClipText randomText

		.AssertEqual cb.TrimTheTail(CreateObject("htmlfile").parentWindow.ClipboardData.GetData("text")), randomText

	.it "should get text from the clipboard"

End With

'teardown

Set fso = Nothing
