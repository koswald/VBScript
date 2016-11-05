
'Provides various properties to analyze a file's encoding

'Usage example

''With CreateObject("includer")
''    Execute(.read("EncodingAnalyzer"))
''End With
'' 
''With New EncodingAnalyzer.SetFile(WScript.Arguments(0))
''    MsgBox "isUTF16LE: " & .isUTF16LE
''End With
'
'Stackoverflow references: <a href="http://stackoverflow.com/questions/3825390/effective-way-to-find-any-files-encoding"> 1</a>, <a href="http://stackoverflow.com/questions/1410334/filesystemobject-reading-unicode-files"> 2</a>.
'
Class EncodingAnalyzer
    Private fso
    Private file
    Private byte0, byte1, byte2, byte3
    Private scriptName

    Sub Class_Initialize
        Set fso = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next 'if called from a .wsc file, the WScript object is unlikely to be available
            scriptName = WScript.ScriptName
        On Error Goto 0
    End Sub

    Private Sub ResetBytes
        byte0 = 0 : byte1 = 0 : byte2 = 0 : byte3 = 0
    End Sub

    'Function SetFile
    'Parameter: a filespec
    'Returns an object self reference
    'Remark: Required. Specifies the file whose encoding is to be determined.

    Function SetFile(file_)
        file = file_
        If Not fso.FileExists(file) Then
            file = file & ".vbs"
            If Not fso.FileExists(file) Then Err.Raise 1, scriptName, "Couldn't find file " & file
        End If
        GetBytes
        Set SetFile = me
    End Function

    'Function isUTF16LE
    'Returns a boolean
    'Remark: Returns a boolean indicating whether the file specified by SetFile is Unicode "Little Endian."

    Function isUTF16LE
        If byte0 = 255 And byte1 = 254 Then isUTF16LE = True Else isUTF16LE = False
    End Function

    'Function isUTF16BE
    'Returns a boolean
    'Remark: Returns a boolean indicating whether the file specified by SetFile is Unicode "Big Endian."

    Function isUTF16BE
        If byte0 = 254 And byte1 = 255 Then isUTF16BE = True Else isUTF16BE = False
    End Function

    'Function isUTF7
    'Returns a boolean
    'Remark: Returns a boolean indicating whether the file specified by SetFile is UTF7.

    Function isUTF7
        If byte0 = &H2b And byte1 = &H2f And byte2 = &H76 Then isUTF7 = True Else isUTF7 = False
    End Function

    'Function isUTF8
    'Returns a boolean
    'Remark: Returns a boolean indicating whether the file specified by SetFile is UTF8.

    Function isUTF8
        If byte0 = &Hef And byte1 = &Hbb And byte2 = &Hbf Then isUTF8 = True Else isUTF8 = False
    End Function

    'Function isUTF32
    'Returns a boolean
    'Remark: Returns a boolean indicating whether the file specified by SetFile is UTF32.

    Function isUTF32
        If byte0 = 0 & byte1 = 0 & byte2 = &Hfe & byte3 = &Hff Then isUTF32 = True Else isUTF32 = False
    End Function

    Sub GetBytes
        ResetBytes
        Dim i, stream
        Set stream = fso.OpenTextFile(file, 1, False)
        For i = 0 To 3
            Execute("byte" & i & " = Asc(stream.Read(1))")
        Next
        stream.Close
        Set stream = Nothing
    End Sub

    Sub Class_Terminate
        Set fso = Nothing
    End Sub
End Class

