
'test VBSNotifyIcon.dll

Option Explicit
Dim test : test = _
"Check system tray icon:" & vbLf & _
"Hover over icon - should display tooltip" & vbLf & _
"Left click icon - should show self-destructing message" & vbLf & _
"Right click icon - should show context menu" & vbLf & _
"Click this OK button - icon should disappear"

Dim ni : Set ni = CreateObject("VBSNotifyIcon")

ni.debug = True
ni.Text = "VBScript system tray icon test"
ni.SetIconByDllFile "%SystemRoot%\System32\msdt.exe", 0
'ni.BalloonTipLifetime = 20000 'deprecated
ni.BalloonTipTitle = "VBSNotifyIcon test"
ni.BalloonTipText = "This message will self-destruct"
ni.SetBalloonTipIcon ni.ToolTipIcon.Info 'Error, Info, None, Warning
ni.Visible = True
MsgBox test, vbSystemModal + vbInformation, WScript.ScriptName
ni.Dispose
Set ni = Nothing
