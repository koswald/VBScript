
'Regular Expression functions - a work in progress
'
'Usage example
'<pre>  With CreateObject("VBScripting.Includer")<br />      Execute .read("RegExFunctions")<br />  End With<br />  <br />  Dim reg : Set reg = New RegExFunctions<br />  reg.SetTestString "'Method SetSomething"<br />  reg.SetPattern "(M).*(od).*(tS)"<br />  <br />  Dim s, submatch, subs : s = ""<br />  Set subs = reg.GetSubMatches<br />  <br />  For Each submatch In subs<br />      s = s & " " & submatch<br />  Next<br />  MsgBox s 'M od tS </pre>
'
Class RegExFunctions

    Private oRE, Match, Matches, testString, class_, v, reader

    Sub Class_Initialize
        With CreateObject("VBScripting.Includer")
            Execute .read("VBSValidator")
        End With
        Set oRE = New RegExp
        Set v = New VBSValidator
        class_ = "RegExFunctions"
        SetPattern ""
        SetTestString ""
        SetIgnoreCase False
    End Sub
    
    'Function Pattern
    'Parameter: wildcard
    'Returns: a regex expression
    'Remark: Returns a regex expression equivalent to the specified wildcard expression(s). Delimit multiple wildcards with &#124;
    Function Pattern(wildcard)
        'See docs\algorithm\ReadMe.md for more comments
        Dim i, arrwp, wp : wp = wildcard '=> wildcard-to-pattern
        'remove whitespace from ends of delimited strings
        wp = Split(wp, "|")
        For i = 0 To UBound(wp)
            wp(i) = Trim(wp(i))
        Next
        wp = Join(wp, "|")
        'raise errors on bad characters
        Dim raise : raise = Array("\", "/", ":", """", "<", ">")
        For i = 0 To UBound(raise)
            If InStr(wp, raise(i)) Then Err.Raise 1,, "A wildcard expression can't contain these: " & Join(raise, " ") & " (In this case " & raise(i) & ")"
        Next
        wp = Replace(wp, ".", "\.")
        wp = Replace(wp, "*", ".*")
        Dim escape : escape = Array("(", ")", "$", "+", "[", "^", "{")
        For i = 0 To UBound(escape)
            wp = Replace(wp, escape(i), "\" & escape(i))
        Next
        wp = Replace(wp, "?", ".{1}")
        arrwp = Split(wp, "|")
        For i = 0 To UBound(arrwp)
            arrwp(i) = "^" & arrwp(i) & "$"
        Next
        Pattern = Join(arrwp, "|")
    End Function

    'Property re
    'Returns an object reference
    'Remark: Returns a reference to the RegExp object instance
    Property Get re : Set re = oRE : End Property

    'Method SetPattern
    'Parameter: a regex pattern
    'Remark: Required before calling FirstMatch or GetSubMatches. Sets the pattern of the RegExp object instance
    Sub SetPattern(pPattern) : oRE.pattern = pPattern : End Sub

    'Method SetTestString
    'Parameter: a string
    'Remark: Required before calling FirstMatch or GetSubMatches. Specifies the string against which the regex pattern will be tested.
    Sub SetTestString(pString) : testString = pString : End Sub

    'Method SetIgnoreCase
    'Parameter: a boolean
    'Remark: Optional. Specifies whether the regex object will ignore case. Default is False.
    Sub SetIgnoreCase(pBool) : oRE.IgnoreCase = v.EnsureBoolean(pBool) : End Sub

    'Method SetGlobal
    'Parameter: a boolean
    'Remark: Optional. Specifies whether the pattern should match all occurrences in the search string or just the first one. Default is False.
    Sub SetGlobal(pBool) : oRE.Global = v.EnsureBoolean(pBool) : End Sub

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
        If Err Then GetSubMatches = Nothing
    End Property

    Private Sub EnsureInitialized
        Dim funct : funct = class_ & ".EnsureInitialized"
        If "" = oRE.pattern Then Err.Raise 2, funct, "RegEx pattern was not set: use SetPattern()"
        If "" = testString Then Err.Raise 3, funct, "RegEx test string was not set: use SetTestString()"
    End Sub

    'Function FirstMatch
    'Returns a string
    'Remark: Regarding the string specified by SetTestString, returns the first substring in the string that matches the regex pattern specified by SetPattern.
    Function FirstMatch
        EnsureInitialized
        FirstMatch = ""
        Set Matches = oRE.Execute(testString)
        For Each Match in Matches
            FirstMatch = Match.Value
            Exit For
        Next
    End Function

End Class
