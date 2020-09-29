' Test fixture for ..\VBSApp.spec.vbs

' This file contains the core VBScript code under test and is referenced by both the .hta and the .wsf fixture files.

With New VBSAppTester
    .Run
End With

Class VBSAppTester

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
            Else stream.WriteLine Err ' write the error number (0)
            End If
        On Error Goto 0

        'output the actual sleep duration
        stopwatch.Reset
        app.Sleep app.GetArg(2)
        stream.WriteLine stopwatch.Split
    End Sub

    Private app, stopwatch, fso, stream

    Sub Class_Initialize
        Set app = CreateObject("VBScripting.VBSApp")
        If "HTMLDocument" = TypeName(document) Then
            app.Init document
        Else app.Init WScript
        End If
        With CreateObject("VBScripting.Includer")
            Dim base
            Execute .read("..\spec\VBSApp.spec.config")
            Execute .read("VBSStopwatch")
        End With
        Set stopwatch = New VBSStopwatch
        Set fso = CreateObject("Scripting.FileSystemObject")
        Const ForWriting = 2
        Const CreateNew = True
        outFile = fso.GetAbsolutePathName( base & app.GetArg(3) & "Out.txt" )
        Set stream = fso.OpenTextFile( outFile, ForWriting, CreateNew )

        Dim outFile
    End Sub

    Sub Class_Terminate
        stream.Close
        Set stream = Nothing
        Set fso = Nothing
        app.Quit
    End Sub

End Class
