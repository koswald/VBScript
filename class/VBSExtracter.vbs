
'For extracting a string from a text file, given a regular expression

Class VBSExtracter

    Private file, oRegExFunctions, oStreamer

    Sub Class_Initialize 'event fires on object instantiation
        With CreateObject("includer") 'get dependencies
            Execute(.read("RegExFunctions"))
            Execute(.read("TextStreamer"))
        End With
        Set oRegExFunctions = New RegExFunctions
        Set oStreamer = New TextStreamer
        SetFile(vbNull)
        SetPattern(vbNull)
        SetTestString(vbNull)
    End Sub

    'Method SetPattern
    'Parameter: a regex pattern
    'Remark: Required. Specifies the text to be extracted. Non-regex expressions containing any of the regex special characters <strong>(  )  .  $  +  [  ?  \  ^  {  |</strong> must preceed the special character with a <strong>\</strong>
    Sub SetPattern(pStr) : re.SetPattern(pStr) : End Sub

    'Method SetFile
    'Parameter: filespec
    'Remark: Required. Specifies the file to extract text from.
    Sub SetFile(pFile) : file = fs.Expand(pFile) : End Sub

    Private Sub SetTestString(pStr) : re.SetTestString(pStr) : End Sub

    'Method SetIgnoreCase
    'Parameter: a boolean
    'Remark: Set whether to ignore case when matching text. Default is False.
    Sub SetIgnoreCase(pBool) : re.SetIgnoreCase(pBool) : End Sub

    ''Method SetMultiline
    ''Parameter: a boolean
    ''Remark: Set whether the RegExp object multiline property is set.
    'Sub SetMultiline(pBool) : re.Multiline = True : End Sub    

    'Sub SetGlobal(pBool) : re.Global(pBool) : End Sub

    'wrap included objects for convenience
    Property Get re : Set re = oRegExFunctions : End Property
    Property Get streamer : Set streamer = ts : End Property
    Property Get ts : Set ts = oStreamer : End Property
    Property Get fs : Set fs = ts.fs : End Property
    Property Get native : Set native = n : End Property
    Property Get n : Set n = ts.n : End Property
    Property Get shell : Set shell = sh : End Property
    Property Get sh : Set sh = ts.sh : End Property
    Property Get fso : Set fso = ts.fso : End Property
    Property Get args : Set args = a : End Property
    Property Get a : Set a = ts.a : End Property

    Private Sub EnsureInitialized
        Dim funct : funct = "VBSExtracter.EnsureInitialized"
        If vbNull = file Then Err.Raise 1, funct, "File to extract text from was not specified. Use SetFile()."
        If vbNull = re.re.Pattern Then Err.Raise 2, funct, "RegEx test pattern was never set."
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
        Dim match : match = re.FirstMatch
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
            match = re.FirstMatch
            If Len(match) Then 'we are at the correct line, and we now have what we need from the file
                Extract0 = match 'return the match value
                Exit Do 'we have what we need, so there is no need to read the rest of the file
            End If
        Loop
        inputStream.Close
        Set inputStream = Nothing
    End Function

End Class
