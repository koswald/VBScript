'test fixture for ..\VBSApp.spec.vbs

'contains the statements included in,
'and common to, the .hta and .wsf
'fixture files

With New Test
    .Run
End With

Class Test

    Sub Run

        stream.WriteLine app.GetArg(1)
        stream.WriteLine app.GetArgsString
        stream.WriteLine app.GetArgsCount
        stream.WriteLine app.GetFullName
        stream.WriteLine app.GetFileName
        stream.WriteLine app.GetBaseName
        stream.WriteLine app.GetExtensionName
        stream.WriteLine app.GetParentFolderName
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
        Set app = CreateObject("VBScripting.VBSApp")
        If "HTMLDocument" = TypeName(document) Then
            app.Init document
        Else app.Init WScript
        End If
        With CreateObject("VBScripting.Includer")
            Dim base
            Execute .read("..\spec\VBSApp.spec.config")
            Execute .read("VBSTimer")
        End With
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
