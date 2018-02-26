
'Functions related to VBScript command-line arguments
'
Class VBSArguments

    'Property GetArgumentsString
    'Returns a string containing all command-line arguments
    'Remark: For use when restarting a script, in order to retain the original arguments. Each argument is wrapped wih quotes, which are stripped off as they are read back in. The return string has a leading space, by design, unless there are no arguments
    Public Default Property Get GetArgumentsString
        Dim s, arg : s = ""
        For Each arg In WScript.Arguments
            s = s & " """ & arg & """"
        Next
        GetArgumentsString = s
    End Property

End Class
