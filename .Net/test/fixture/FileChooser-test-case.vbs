L = vbLf : T = "       " : LT = L & T
expectedOutcome = _
"Launch FileChooser-test.vbs." & LT & _
    "This message should appear on top" & LT & _
    "of other windows." & L & _
"Move this message to one side of the screen," & L & _
"away from the Browse For File dialog," & L & _
"and away from the center of the screen." & LT & _
    "A Browse For File dialog should have appeared." & LT & _
    """All files (*.*)"" should be selected." & LT & _
    """RTF files (*.rtf)"" should be another option." & L & _
"Select a file and click Open" & L & _
"(or double-click a file)" & LT & _
    "A message box should display the filespec" & LT & _
    "of the selected file." & L & _
"Click OK to continue or Cancel to end the test." & LT & _
    "Another browse for a file dialog should appear." & L & _
"Select two or more files and click Open." & LT & _
    "A message box should display the filespecs" & LT & _
    "of the selected files."

MsgBox expectedOutcome, vbSystemModal + vbInformation, WScript.ScriptName

