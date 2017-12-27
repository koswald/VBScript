
'Add the VBScripting source to the Application log

Option Explicit : Initialize

Const source = "VBScripting"

Call Main

Sub Main

    'create the source
    Dim result : Set result = va.CreateEventSource(source)

    'show the result
    If Not quiet Then
        MsgBox "Result: " & result.Message, vbInformation, result.Result
    End If

    ReleaseObjectMemory
End Sub

Sub ReleaseObjectMemory
    Set va = Nothing
    Set sh = Nothing
End Sub

Dim va, sh, format
Dim quiet

Sub Initialize
    Set va = CreateObject("VBScripting.Admin")
    Set sh = CreateObject("WScript.Shell")
    With CreateObject("includer")
        Execute .read("StringFormatter")
    End With
    Set format = New StringFormatter

    'check for /quiet flag on the command line
    quiet = False
    Dim args : Set args = WScript.Arguments
    If args.Count Then
        Dim i
        For i = 0 To args.Count - 1
            If "/quiet" = LCase(args.item(i)) Then quiet = True
        Next
    End If

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
