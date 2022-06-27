' Test fixture for ..\VBSApp.spec.vbs

' This file contains the core VBScript code under test and is referenced by both the .hta and the .wsf fixture files.

With New VBSAppTester
    .Run "New object"
    .Run "COM object"
End With

Class VBSAppTester
    Private app 'the VBSApp object under test
    Private fso 'Scripting.FileSystemObject object
    Private incl 'VBScripting.Includer object
    Private sh 'WScript.Shell object
    Private stream 'text stream for writing to the output file
    Private stopwatch 'a project object

    Sub Class_Initialize
        Dim outFile 'filespec of the output file
        Const ForWriting = 2, CreateNew = True 'for OpenTextFile

        Set fso = CreateObject( "Scripting.FileSystemObject" )
        Set sh = CreateObject( "WScript.Shell" )
        Set incl = CreateObject( "VBScripting.Includer" )
        Set app = CreateObject( "VBScripting.VBSApp" )
        If "HTMLDocument" = TypeName(document) Then
            app.Init document
        Else app.Init WScript
        End If
        outFile = Expand( "%AppData%\VBScripting\VBSApp." & app.GetArg(3) & "Out.txt" )
        Set stream = fso.OpenTextFile( outFile, ForWriting, CreateNew )
        Execute incl.Read( "VBSStopwatch" )
        Set stopwatch = New VBSStopwatch
    End Sub

    Sub Run( instantiationMethod )
        Set app = Nothing
        Select Case instantiationMethod

        Case "COM object"
            Set app = CreateObject( "VBScripting.VBSApp" )
            If "HTMLDocument" = TypeName( document ) Then
                app.Init document
            Else app.Init WScript
            End If

        Case "New object"
            Execute incl.Read( "VBSApp" )
            Set app = New VBSApp
        End Select

        stream.WriteLine app.GetArg(1)
        stream.WriteLine app.GetArgsString
        stream.WriteLine app.GetArgsCount
        stream.WriteLine app.GetFullName
        stream.WriteLine app.GetFileName
        stream.WriteLine app.GetBaseName
        stream.WriteLine app.GetExtensionName
        stream.WriteLine app.GetParentFolderName
        stream.WriteLine app.GetExe
        app.RUArgsTest = True
        app.RestartUsing "wscript.exe", app.DoExit, app.DoNotElevate
        stream.WriteLine app.RUArgs

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

    Function Expand( str )
        Expand = sh.ExpandEnvironmentStrings( str )
    End Function

    Sub Class_Terminate
        stream.Close
        Set stream = Nothing
        Set fso = Nothing
        Set sh = Nothing
        Set incl = Nothing
        Set stopwatch = Nothing
        app.Quit
    End Sub

End Class
