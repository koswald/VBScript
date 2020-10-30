' Test fixture for ..\VBSApp.spec.vbs

' This file contains the core VBScript code under test and is referenced by both the .hta and the .wsf fixture files.

With New VBSAppTester
    .Run
End With

Class VBSAppTester

    Sub Run

        If Not "TextStream" = TypeName(stream) Then
            ' Failed to create the output file in Sub Class_Initialize.
            errMsg = "Running the VBSApp integration test from %ProgramFiles% requires elevated privileges."
            sh.PopUp errMsg, timeout, app.GetFileName, vbInformation
            app.Quit
        End If

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

        Const timeout = 20 ' seconds; 0 => indefinite
        Dim errMsg
    End Sub

    Private fso, sh, stream ' Windows-native objects
    Private app, stopwatch ' project objects

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
        On Error Resume Next
            Set stream = fso.OpenTextFile( outFile, ForWriting, CreateNew )
        On Error Goto 0
        Set sh = CreateObject("WScript.Shell")

        Dim outFile
    End Sub

    Sub Class_Terminate
        stream.Close
        Set stream = Nothing
        Set fso = Nothing
        app.Quit
    End Sub

End Class
