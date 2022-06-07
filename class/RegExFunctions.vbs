'Regular Expression functions - a work in progress
'
'Usage example
'<pre>  With CreateObject( "VBScripting.Includer" )<br />      Execute .Read( "RegExFunctions" )<br />  End With<br />  <br />  Dim reg : Set reg = New RegExFunctions<br />  reg.SetTestString "'Method SetSomething"<br />  reg.SetPattern "(M).*(od).*(tS)"<br />  <br />  Dim s, submatch, subs : s = ""<br />  Set subs = reg.GetSubMatches<br />  <br />  For Each submatch In subs<br />      s = s & " " & submatch<br />  Next<br />  MsgBox s 'M od tS </pre>
'
Class RegExFunctions

    Private rex 'RegExp object
    Private Match
    Private Matches
    Private testString 'string against which the regex pattern will be tested
    Private v 'VBSValidator object

    Sub Class_Initialize
        With CreateObject( "VBScripting.Includer" )
            Execute .Read( "VBSValidator" )
        End With
        Set rex = New RegExp
        Set v = New VBSValidator
        SetPattern ""
        SetTestString ""
        SetIgnoreCase False
    End Sub

    'Function Pattern
    'Parameter: wildcard
    'Returns: a regex expression
    'Remark: Returns a regular expression equivalent to the specified wildcard expression(s). Delimit multiple wildcards with a vertical bar ( &#124; ). See <a href=https://github.com/koswald/VBScript/blob/master/docs/algorithm/ReadMe.md target=_blank> algorithm/ReadMe.md</a> for more comments.
    Function Pattern( wildcard )
        Dim arr 'array
        Dim i 'integer
        Dim str 'string
        Dim raise 'array: raise an error on finding these characters
        Dim escape 'array: escape these characters

        'remove whitespace from ends of delimited strings
        arr = Split( wildcard, "|" )
        For i = 0 To UBound( arr )
            arr( i ) = Trim( arr( i ))
        Next
        str = Join( arr, "|" )

        'raise an error on finding certain characters
        raise = Array( "\", "/", ":", """", "<", ">" )
        For i = 0 To UBound( raise )
            If InStr( str, raise( i )) Then
                Err.Raise 5,, "A wildcard expression can't contain these: " & Join( raise ) & " ( In this case " & raise( i ) & " )."
            End If
        Next
        
        'replace . and * with regex equivalents
        str = Replace( str, ".", "\." )
        str = Replace( str, "*", ".*" )
        
        'escape certain characters
        escape = Array( "(", ")", "$", "+", "[", "^", "{" )
        For i = 0 To UBound( escape )
            str = Replace( str, escape( i ), "\" & escape( i ))
        Next
        
        'replace ? with regex equivalent
        str = Replace( str, "?", ".{1}" )
        
       'return value
        arr = Split( str, "|" )
        For i = 0 To UBound( arr )
            arr( i ) = "^" & arr( i ) & "$"
        Next
        Pattern = Join( arr, "|" )
    End Function

    'Property re
    'Returns an object reference
    'Remark: Returns a reference to the RegExp object instance.
    Property Get re
        Set re = rex
    End Property

    'Method SetPattern
    'Parameter: a regex pattern
    'Remark: Required before calling FirstMatch or GetSubMatches. Sets the pattern of the RegExp object instance.
    Sub SetPattern( newPattern )
        rex.pattern = newPattern
    End Sub

    'Method SetTestString
    'Parameter: a string
    'Remark: Required before calling FirstMatch or GetSubMatches. Specifies the string against which the regex pattern will be tested.
    Sub SetTestString( newValue )
        testString = newValue
    End Sub

    'Method SetIgnoreCase
    'Parameter: a boolean
    'Remark: Optional. Specifies whether the regex object will ignore case. Default is False.
    Sub SetIgnoreCase( newValue )
        rex.IgnoreCase = v.EnsureBoolean( newValue )
    End Sub

    'Method SetGlobal
    'Parameter: a boolean
    'Remark: Optional. Specifies whether the pattern should match all occurrences in the search string or just the first one. Default is False.
    Sub SetGlobal( newValue )
        rex.Global = v.EnsureBoolean( newValue )
    End Sub

    'Property GetSubMatches
    'Returns an object
    'Remark: Returns the RegExp SubMatches object for the specified pattern and test string. The matches can be accessed with a For Each loop. See general usage comments. Work in progress. You must handle errors in case there are no matches.
    'TODO: This is unwieldy. It would be better to return an array of strings, because a zero-length array could be more easily handled (?).
    Property Get GetSubMatches
        On Error Resume Next
        Dim oMatch, oMatches
        Set oMatches = re.Execute(testString)
        Set oMatch = oMatches(0)
        Set GetSubMatches = oMatch.SubMatches
        If Err Then
            Set GetSubMatches = Nothing
        End If
    End Property

    Private Sub EnsureInitialized
        If "" = rex.pattern Then Err.Raise 449,, "RegEx pattern was not set: use SetPattern()"
        If "" = testString Then Err.Raise 449,, "RegEx test string was not set: use SetTestString()"
    End Sub

    'Function FirstMatch
    'Returns a string
    'Remark: Regarding the string specified by SetTestString, returns the first substring in the string that matches the regex pattern specified by SetPattern.
    Function FirstMatch
        EnsureInitialized
        FirstMatch = ""
        Set Matches = rex.Execute( testString )
        For Each Match in Matches
            FirstMatch = Match.Value
            Exit For
        Next
    End Function

End Class
