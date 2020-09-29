' Test for NamedArguments.vbs

If vbCancel = MsgBox("Expected outcome: After this message, if the test is continued by clicking OK, another message should appear showing the script syntax/usage, because /Folder: was not specified, and it is a required argument.", vbOKCancel + vbInformation + vbSystemModal, WScript.ScriptName) Then WScript.Quit

Dim sh : Set sh = CreateObject("WScript.Shell")
sh.Run "NamedArguments.wsf",, synchronous

sh.Run "NamedArguments.wsf /Folder:. /ExpectedOutcome:""A listing of the default file types, .vbs and .wsf, in the current folder.""",, synchronous

sh.Run "NamedArguments.wsf /Folder:. /FileTypes:"" md | wsf "" /ExpectedOutcome:""A listing of .md and .wsf files in the current folder.""",, synchronous

sh.Run "NamedArguments.wsf /Folder:.. /FileTypes:"" md | hta | vbs "" & /ExpectedOutcome:""A listing of .md, .hta, and .vbs files in the parent folder.""",, synchronous

MsgBox "Expected outcome: After this message an error should be received that the indicated folder does not exist", vbInformation + vbSystemModal, WScript.ScriptName

Set sa = CreateObject("Shell.Application")
sa.MinimizeAll

sh.Run "NamedArguments.wsf /Folder:.\NoSuchFolder",, synchronous

sa.UndoMinimizeAll
Set sa = Nothing
Set sh = Nothing

Const synchronous = True