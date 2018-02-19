' elevate privileges
With WScript.Arguments
    If .Count = 0 Then
        Err.Raise 1,, "Expected a command line argument: the file to open with elevated privileges."
    End If
    For i = 0 To .Count - 1
        If Instr(.item(i), " ") Then
            args = args & " """ & .item(i) & """"
        Else args = args & " " & .item(i)
        End If
    Next
End With
With CreateObject("Shell.Application")
    .ShellExecute "cmd", "/c start " & args,, "runas"
End With
