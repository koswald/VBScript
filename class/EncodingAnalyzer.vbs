
'Provides various properties to analyze a file's encoding

'Usage example

''With CreateObject("VBScripting.Includer")
''    Execute .read("EncodingAnalyzer")
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
    Private sh
    Private file, fileHasBeenValidated
    Private byte0, byte1, byte2, byte3
    Private scriptName

    Sub Class_Initialize
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set sh = CreateObject("WScript.Shell")
        fileHasBeenValidated = False
    End Sub

    Private Sub ResetBytes
        byte0 = 0 : byte1 = 0 : byte2 = 0 : byte3 = 0
    End Sub

    'Function SetFile
    'Parameter: a filespec
    'Returns an object self reference
    'Remark: Required. Specifies the file whose encoding is to be determined. Relative paths are permitted, relative to the current directory.
    Function SetFile(file_)
        file = file_
        fileHasBeenValidated = False
        ValidateFile
        GetBytes
        Set SetFile = me
    End Function

    Private Sub ValidateFile
        If fileHasBeenValidated Then Exit Sub
        If Not fso.FileExists(file) Then
            file = file & ".vbs"
            If Not fso.FileExists(file) Then Err.Raise 1,, "The EncodingAnalyzer couldn't find file " & file
        End If
        fileHasBeenValidated = True
    End Sub

    'Function isUTF16LE
    'Returns a boolean
    'Remark: Returns a boolean indicating whether the file specified by SetFile is Unicode Little Endian, <strong> aka Unicode</strong>.
    Function isUTF16LE
        ValidateFile
        If byte0 = &Hff And byte1 = &Hfe Then isUTF16LE = True Else isUTF16LE = False
    End Function

    'Function isUTF16BE
    'Returns a boolean
    'Remark: Returns a boolean indicating whether the file specified by SetFile is Unicode Big Endian.
    Function isUTF16BE
        ValidateFile
        If byte0 = &Hfe And byte1 = &Hff Then isUTF16BE = True Else isUTF16BE = False
    End Function

    'Function isUTF7
    'Returns a boolean
    'Remark: Returns a boolean indicating whether the file specified by SetFile is UTF7.
    Function isUTF7
        ValidateFile
        If byte0 = &H2b And byte1 = &H2f And byte2 = &H76 Then isUTF7 = True Else isUTF7 = False
    End Function

    'Function isUTF8
    'Returns a boolean
    'Remark: Returns a boolean indicating whether the file specified by SetFile is UTF8.
    Function isUTF8
        ValidateFile
        If byte0 = &Hef And byte1 = &Hbb And byte2 = &Hbf Then isUTF8 = True Else isUTF8 = False
    End Function

    'Function isUTF32
    'Returns a boolean
    'Remark: Returns a boolean indicating whether the file specified by SetFile is UTF32.
    Function isUTF32
        ValidateFile
        If byte0 = 0 And byte1 = 0 And byte2 = &Hfe And byte3 = &Hff Then isUTF32 = True Else isUTF32 = False
    End Function

    'Function isAscii
    'Returns a boolean
    'Remark: Returns a boolean indicating whether the file specified by SetFile is Ascii.
    Function isAscii
        ValidateFile
        isAscii = True
        If isUTF16LE Then isAscii = False : Exit Function
        If isUTF16BE Then isAscii = False : Exit Function
        If isUTF7 Then isAscii = False : Exit Function
        If isUTF8 Then isAscii = False : Exit Function
        If isUTF32 Then isAscii = False
    End Function

    'Function GetType
    'Returns a string
    'Remark: Returns one of the following strings according the format of the file set by SetFile: Ascii, UTF16LE, UTF16BE, UTF7, UTF8, UTF32.
    Function GetType
        ValidateFile
        GetType = "Ascii"
        If isUTF16LE Then GetType = "UTF16LE" : Exit Function
        If isUTF16BE Then GetType = "UTF16BE" : Exit Function
        If isUTF7 Then GetType = "UTF7" : Exit Function
        If isUTF8 Then GetType = "UTF8" : Exit Function
        If isUTF32 Then GetType = "UTF32"
    End Function

    'Read the first four bytes from the specified file:
    'prepare for analysis
    Private Sub GetBytes
        ValidateFile
        ResetBytes
        Dim i, stream
        Set stream = fso.OpenTextFile(file, 1, False)
        For i = 0 To 3
            Execute("byte" & i & " = Asc(stream.Read(1))")
        Next
        stream.Close
        Set stream = Nothing
    End Sub

    'Function GetCurrentDirectory
    'Returns a folder
    'Remarks: Returns the current directory
    Function GetCurrentDirectory
        GetCurrentDirectory = sh.CurrentDirectory
    End Function

    'Method SetCurrentDirectory
    'Parameter: a folder
    'Remarks: Sets the current directory.
    Sub SetCurrentDirectory(dir)
        sh.CurrentDirectory = dir
    End Sub

    'Function GetByte
    'Parameter: BOM byte number
    'Returns an integer
    'Remark: Returns the Ascii value, 0 to 255, of the byte specified. The parameter must be an integer: one of 0, 1, 2, or 3. These represent the first four bytes in the file, the Byte Order Mark (BOM).
    Function GetByte(i)
        ValidateFile
        Execute("GetByte = byte" & i)
    End Function

    Sub Class_Terminate
        Set fso = Nothing
        Set sh = Nothing
    End Sub
End Class

