
'For extracting a string from a text file, given a regular expression

Class VBSExtracter

    Private file
    Private ref, streamer, fs, fso

    Sub Class_Initialize 'event fires on object instantiation
        With CreateObject("VBScripting.Includer") 'get dependencies
            Execute .read("RegExFunctions")
            Execute .read("TextStreamer")
            Execute .read("VBSFileSystem")
        End With
        Set ref = New RegExFunctions
        Set fs = New VBSFileSystem
        Set streamer = New TextStreamer
        Set fso = CreateObject("Scripting.FileSystemObject")
        SetFile(vbNull)
        SetPattern(vbNull)
        SetTestString(vbNull)
    End Sub

    'Method SetPattern
    'Parameter: a regex pattern
    'Remark: Required. Specifies the text to be extracted. Non-regex expressions containing any of the regex special characters <strong>(  )  .  $  +  [  ?  \  ^  {  |</strong> must preceed the special character with a <strong>\</strong>
    Sub SetPattern(pStr) : ref.SetPattern(pStr) : End Sub

    'Method SetFile
    'Parameter: filespec
    'Remark: Required. Specifies the file to extract text from.
    Sub SetFile(pFile) : file = fs.Expand(pFile) : End Sub

    Private Sub SetTestString(pStr) : ref.SetTestString(pStr) : End Sub

    'Method SetIgnoreCase
    'Parameter: a boolean
    'Remark: Set whether to ignore case when matching text. Default is False.
    Sub SetIgnoreCase(pBool) : ref.SetIgnoreCase(pBool) : End Sub

    ''Method SetMultiline
    ''Parameter: a boolean
    ''Remark: Set whether the RegExp object multiline property is set.
    'Sub SetMultiline(pBool) : ref.Multiline = True : End Sub    

    'Sub SetGlobal(pBool) : ref.Global(pBool) : End Sub

    Private Sub EnsureInitialized
        Dim funct : funct = "VBSExtracter.EnsureInitialized"
        If vbNull = file Then Err.Raise 1, funct, "File to extract text from was not specified. Use SetFile()."
        If vbNull = ref.re.Pattern Then Err.Raise 2, funct, "RegEx test pattern was never set."
        If Not fso.FileExists(file) Then Err.Raise 3, funct, "Couldn't find the file to extract text from, " & vbLf & vbTab & file
    End Sub

    'Property Extract
    'Returns a string
    'Remark: Returns the first string that matches the specified regex pattern. Returns an empty string if there is no match. Before calling this method, you must specify the file and the pattern: see SetPattern and SetFile.
    Function Extract
        Extract = ""
        EnsureInitialized
        streamer.SetFile(file) 'set the streamer to use the file specified
        streamer.SetForReading 'set the streamer for reading
        Dim inputStream : Set inputStream = streamer.Open 'open the file as a text stream
        SetTestString(inputStream.ReadAll)
        Dim match : match = ref.FirstMatch
        If Len(match) Then
            Extract = match
        End If
        inputStream.Close
        Set inputStream = Nothing
    End Function

    'Property Extract0
    'Returns a string
    'Remark: Deprecated for not spanning multiple lines. Formerly named Extract. Returns the string that matches the specified regex pattern. Returns an empty string if there is no match. Before calling this method, you must specify the file and the pattern: see SetPattern and SetFile.
    Private Function Extract0
        Extract0 = ""
        EnsureInitialized
        streamer.SetFile(file) 'set the streamer to use the file specified
        streamer.SetForReading 'set the streamer for reading
        Dim inputStream : Set inputStream = streamer.Open 'open the file as a text stream
        Dim match
        Do Until inputStream.AtEndOfStream 'or until the Exit Do statement is reached
            SetTestString(inputStream.ReadLine)
            match = ref.FirstMatch
            If Len(match) Then 'we are at the correct line, and we now have what we need from the file
                Extract0 = match 'return the match value
                Exit Do 'we have what we need, so there is no need to read the rest of the file
            End If
        Loop
        inputStream.Close
        Set inputStream = Nothing
    End Function

    Sub Class_Terminate
        Set fso = Nothing
    End Sub

End Class
