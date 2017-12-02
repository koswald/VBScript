
'Add the VBScripting source to the Applications log

Option Explicit : Initialize

Const source = "VBScripting"
Const configFile = "../../spec/dll/Admin.spec.config"

Call Main
Sub Main

    'create the source
    va.CreateEventSource source

    'option to open the .config file
    OpenTheConfigFile

    ReleaseObjectMemory
End Sub

Sub ReleaseObjectMemory
    Set va = Nothing
    Set sh = Nothing
End Sub

Sub OpenTheConfigFile
    Dim msg : msg = "If the source was created " & _
        "successfully, then change the source in the " & _
        "file " & configFile & " to " & source & "."
    Dim mode : mode = vbInformation + vbSystemModal + vbOKCancel
    Dim caption : caption = WScript.ScriptName
    If vbOK = MsgBox(msg, mode, caption) Then
        sh.Run "notepad " & configFile
    End If
End Sub

Dim va, sh

Sub Initialize
    Set va = CreateObject("VBScripting.Admin")
    Set sh = CreateObject("WScript.Shell")

    If va.PrivilegesAreElevated Then Exit Sub

    'elevate privileges
    ReleaseObjectMemory
    With CreateObject("includer")
        Execute .read("VBSApp")
    End With
    With New VBSApp
        .SetUserInteractive False
        .RestartWith "wscript", "/c", True
    End With
End Sub
