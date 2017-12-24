
''' Fixture for the TestingFramework class.
''' Shows a SendKeys warning message.

If WScript.Arguments.Count Then
    this = " " & WScript.Arguments.item(0)
Else this = ""
End If

warning = "Don't make mouse clicks or key presses while the" & _
    this & " test is in progress!" & vbLf & vbLf & _
    "To cancel the test, close the console window."
mode = vbSystemModal + vbExclamation
caption = "Warning"

MsgBox warning, mode, caption