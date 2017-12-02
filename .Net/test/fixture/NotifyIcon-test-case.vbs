
'show test case info

'intended to be called automatically

L = vbLf : T = "       " : LT = L & T : LTT = LT & T

test = _
"Launch NotifyIcon-test.vbs" & LT & _
    "This message should appear" & L & _
"Hover over the system tray icon." & LT & _
    "A tooltip should appear" & L & _
"Right click the icon" & LT & _
    "A context menu should appear" & L & _
"Left click the icon" & LT & _
    "The same context menu should appear" & L & _
"Select Show balloon tip from the context menu" & LT & _
    "A notification should appear" & L & _
"Select Exit from the context menu" & LT & _
    "The icon and this dialog should disappear"

MsgBox test, vbSystemModal + vbInformation, WScript.ScriptName
