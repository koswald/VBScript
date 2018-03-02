
'Default property Privileged returns True if the calling script has elevated privileges.

'Usage example
'<pre> With CreateObject("VBScripting.Includer") <br />     Execute .read("PrivilegeChecker") <br /> End With <br /> Dim pc : Set pc = New PrivilegeChecker <br /> If pc Then <br />     WScript.Echo "Privileges are elevated" <br /> Else <br />     WScript.Echo "Privileges are not elevated" <br /> End If </pre>
'
'Reference: <a href="http://stackoverflow.com/questions/4051883/batch-script-how-to-check-for-admin-rights/21295806"> stackoverflow.com</a>
'
'''The fsutil technique works with Windows XP thru 10.

Class PrivilegeChecker

    'Function Privileged
    'Returns a boolean
    'Remark: Returns True if the calling script is running with elevated privileges, False if not. Privileged is the default property.
    Public Default Function Privileged

        Dim sh : Set sh = CreateObject("WScript.Shell")
        Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim privileged_, unprivileged_, undefined_
        privileged_ = "privileged"
        unprivileged_ = "unprivilegd" 'intentionally misspelled for unique search results
        undefined_ = "undefined"
        Privileged = undefined_

        'create a randomly-named .bat file
        Dim tempFile : tempFile = sh.ExpandEnvironmentStrings("%temp%\" & fso.GetTempName & ".bat")
        Dim bf : Set bf = fso.OpenTextFile(tempFile, 2, True) 'create the batch file; open for writing
        bf.WriteLine "@echo off"
        bf.WriteLine "call :isAdmin"
        bf.WriteLine "if %errorlevel% == 0 ("
        bf.WriteLine "echo " & privileged_
        bf.WriteLine ") else ("
        bf.WriteLine "echo " & unprivileged_
        bf.WriteLine ")"
        bf.WriteLine "exit /b"
        bf.WriteLine ":isAdmin"
        bf.WriteLine "fsutil dirty query %systemdrive% >nul"
        bf.WriteLine "exit /b"
        bf.Close
        Set bf = Nothing

        'run the batch file and parse the output
        Dim pipe : Set pipe = sh.Exec("%ComSpec% /c """ & tempFile & """")
        Dim line
        While Not pipe.StdOut.AtEndOfStream
            line = pipe.StdOut.ReadLine
            If InStr(line, privileged_) Then
                Privileged = True
            ElseIf InStr(line, unprivileged_) Then
                Privileged = False
            End If
        Wend

        'cleanup
        Set pipe = Nothing
        fso.DeleteFile(tempFile)
        Set sh = Nothing
        Set fso = Nothing

        'raise an error if privileges are undefined
        If Privileged = undefined_ Then Err.Raise 1,, "The PrivilegeChecker could not determine privileges"
    End Function

End Class
