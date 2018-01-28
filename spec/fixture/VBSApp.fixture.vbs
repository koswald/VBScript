'test fixture for ..\VBSApp.spec.vbs

'contains the statements included in,
'and common to, the .hta and .wsf
'fixture files


With New Test
    .Run
End With

Class Test

    Sub Run

        'output selected command-line argument
        stream.WriteLine app.GetArg(1) 

        'output the command-line string
        stream.WriteLine app.GetArgsString

        'output the argument count
        stream.WriteLine app.GetArgsCount

        'output the filespec
        stream.WriteLine app.GetFullName

        'output the file name
        stream.WriteLine app.GetFileName

        'output the base file name
        stream.WriteLine app.GetBaseName

        'output the file extension name
        stream.WriteLine app.GetExtensionName

        'output the parent folder name
        stream.WriteLine app.GetParentFolderName

        'output the host .exe
        stream.WriteLine app.GetExe

        'attempt to invoke the Sleep method
        On Error Resume Next
            app.Sleep 1
            If Err Then
                stream.WriteLine Err.Description
            Else
                stream.WriteLine Err
            End If
        On Error Goto 0

        'output the actual sleep duration
        tmr.Reset
        app.Sleep app.GetArg(2)
        stream.WriteLine tmr.Split

    End Sub 'Run

    Private app, tmr, fso, stream

    Sub Class_Initialize
        With CreateObject("includer")
            Dim base
            Execute .read("..\spec\VBSApp.spec.config")
            Execute .read("VBSApp")
            Execute .read("VBSTimer")
        End With
        Set app = New VBSApp
        Set tmr = New VBSTimer
        Set fso = CreateObject("Scripting.FileSystemObject")
        Const ForWriting = 2
        Const CreateNew = True
        Set stream = fso.OpenTextFile(base & app.GetArg(3) & "Out.txt", ForWriting, CreateNew)
    End Sub

    Sub Class_Terminate
        stream.Close
        Set stream = Nothing
        Set fso = Nothing
        app.Quit
    End Sub

End Class
