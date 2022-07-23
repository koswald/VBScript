'Manual integration test for the VBScripting.NotifyIcon object and .dll

Option Explicit
Dim ni 'the VBScripting.NotifyIcon object to be tested
Dim sh 'WScript.Shell object
Dim fso 'Scripting.FileSystemObject
Dim testCaseMessage 'wscript process object/Exec method return value: shows a MsgBox with test instructions.
Dim format 'VBScripting.StringFormatter object
Const closedByTester = 1 'testCaseMessage status
Const largeIcon = True, smallIcon = False
Dim icons '1-based array: specifies a collection of icons: includes for each icon a filespec, index, size (boolean), and description.
Dim iconIndex 'integer
Dim shell32_dll 'filespec

Set sh = CreateObject( "WScript.Shell" )
Set fso = CreateObject( "Scripting.FileSystemObject" )
Set format = CreateObject( "VBScripting.StringFormatter")
sh.CurrentDirectory = fso.GetParentFolderName( WScript.ScriptFullName )
Set testCaseMessage = sh.Exec("wscript fixture\NotifyIcon-test-case.vbs")
shell32_dll = "%SystemRoot%\System32\shell32.dll"
icons = Array("" _
    , shell32_dll, 77, smallIcon, "exclamation" _
    , shell32_dll, 43, smallIcon, "gold star" _
    , shell32_dll, 15, smallIcon, "computer" _
)
iconIndex = 1

'Create the object
Set ni = CreateObject( "VBScripting.NotifyIcon" )

'test three ways to set the icon
ni.SetIconByIcoFile "fixture\star.ico"
ni.SetIconByDllFile "%SystemRoot%\System32\shell32.dll", 272, largeIcon
ni.SetIconByDllFile "%SystemRoot%\System32\msdt.exe", 0, largeIcon

'on-hover tooltip
ni.Text = "VBScripting.NotifyIcon" & vbLf & "test"

'balloon tip / notification
ni.BalloonTipTitle = "NotifyIcon test"
ni.BalloonTipText = "Notification message, AKA balloon tip text."
ni.SetBalloonTipIcon ni.ToolTipIcon.Info 'Error, Info, None, Warning

'add menu items / callbacks
ni.AddMenuItem "Show balloon tip", GetRef( "ShowBalloonTip" )
ni.AddMenuItem "Open test file in Notepad", GetRef( "OpenNotepad" )
ni.AddMenuItem "Change the icon", GetRef( "ChangeTheIcon" )
ni.AddMenuItem "E&xit", GetRef( "CloseAndExit" )
ni.Visible = True

'set callback for the balloon tip (notification) click
ni.SetBalloonTipCallback GetRef( "BalloonTipClicked" )

ListenForCallbacks

'Keep the script running in order to listen for callbacks,
'events triggered by the NotifyIcon object.
'This approach should not be used in an .hta file.
Sub ListenForCallbacks
    While True
        WScript.Sleep 200
        If closedByTester = testCaseMessage.Status Then CloseAndExit
    Wend
End Sub

Sub ShowBalloonTip
    ni.ShowBalloonTip
End Sub

Sub BalloonTipClicked
   sh.PopUp "BalloonTip clicked", 20, WScript.ScriptName, vbInformation + vbSystemModal
End Sub

Sub OpenNotepad
    sh.Run "Notepad """ & WScript.ScriptFullName & """"
End Sub

Sub ChangeTheIcon
    If iconIndex > UBound(icons) Then iconIndex = 1
    Dim i : i = iconIndex
    ni.SetIconByDllFile icons(i), icons(i + 1), icons(i + 2)
    'change tooltip to show icon description, filename, and index
    ni.Text = format(Array( _
        "%s: %s index %s", _
        icons(i + 3), fso.GetFileName(icons(i)) , icons(i + 1) _
    ))
    iconIndex = iconIndex + 4
End Sub

Sub CloseAndExit
    testCaseMessage.Terminate
    ni.Dispose
    Set ni = Nothing
    Set sh = Nothing
    Set fso = Nothing
    WScript.Quit
End Sub
