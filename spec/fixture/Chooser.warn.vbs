
'Show a warning while conducting tests that use SendKeys!

warning = "Do not press any keys or make mouse clicks while the Chooser test is in progress!" _
 & vbLf & vbLf & "To cancel, close the console/terminal window."
soundlevel = noisy
prominence = AlwaysOnTop
MsgBox warning, soundLevel + prominence, "Warning!"

Const silent = 0
Const noisy = 48 'may be noisy but shows exclamation icon
Const AlwaysOnTop = 4096
Const QuitePossiblyNotOnTop = 0
