
'Add the VBScripting source to the Application log

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
    Dim msg : msg = format(Array( _
        "If the source was created successfully, " & _
        "then edit %s and change the source to %s.", _
        configFile, source _
    ))
    Dim mode : mode = vbInformation + vbSystemModal + vbOKCancel
    Dim caption : caption = WScript.ScriptName
    If vbOK = MsgBox(msg, mode, caption) Then
        sh.Run "notepad " & configFile
    End If
End Sub

Dim va, sh, format

Sub Initialize
    Set va = CreateObject("VBScripting.Admin")
    Set sh = CreateObject("WScript.Shell")
    With CreateObject("includer")
        Execute .read("StringFormatter")
    End With
    Set format = New StringFormatter

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
