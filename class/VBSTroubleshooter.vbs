Class VBSTroubleshooter

    'Method LogAscii
    'Parameter: a string
    'Remark: Write to the log the Ascii codes for each character in the specified string.
    Sub LogAscii(str)
        Dim i, c, s
        For i = 0 To Len(str) - 1
            s = Right(str, Len(str) - i) 'get substring s, which has the char of interest at its left-most point.
            c = Left(s, 1) 'get char of interest, c
            log format(Array("char: %s, Ascii: %s", c, Asc(c) ))
        Next
    End Sub

    Private log, format

    Sub Class_Initialize
        With CreateObject( "VBScripting.Includer" )
            Execute .Read( "VBSLogger" )
            Execute .Read( "StringFormatter" )
        End With
        Set log = New VBSLogger
        Set format = New StringFormatter
    End Sub
End Class
