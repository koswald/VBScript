'Elevate privileges of the specified file
'Use as a drop target or from the command line.

Option Explicit
With WScript.Arguments
    If .Count = 0 Then
        Err.Raise 1,, "Expected a command line argument: the file to open with elevated privileges."
    End If
    Dim i, args : args = ""
    For i = 0 To .Count - 1
        If InStr(.item(i), " ") Then
            args = args & " """ & .item(i) & """"
        Else args = args & " " & .item(i)
        End If
    Next
End With
Dim sa : Set sa = CreateObject("Shell.Application")
Dim format : Set format = CreateObject("VBScripting.StringFormatter")
Dim sh : Set sh = CreateObject("WScript.Shell")
sa.ShellExecute "cmd", format(Array( _
        "/c cd ""%s"" & start """" %s", _
        sh.CurrentDirectory, args _
    )),, "runas"

Set sa = Nothing : Set format = Nothing : Set sh = Nothing

