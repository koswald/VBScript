
'Regular Expression functions - a work in progress
'
'Usage example
'
'' With CreateObject("includer")
''     Execute(.read("RegExFunctions"))
'' End With
'' 
'' Dim reg : Set reg = New RegExFunctions
'' reg.SetTestString "'Method SetSomething"
'' reg.SetPattern "(M).*(od).*(tS)"
'' 
'' Dim s, submatch, subs : s = ""
'' Set subs = reg.GetSubMatches
'' 
'' For Each submatch In subs
''     s = s & " " & submatch
'' Next
'' MsgBox s 'M od tS
'
Class RegExFunctions

    Private oRE, Match, Matches, testString, class_, v, reader

    Sub Class_Initialize
        With CreateObject("includer") : On Error Resume Next
            Execute(.read("VBSValidator"))
        End With : On Error Goto 0
        Set oRE = New RegExp
        Set v = New VBSValidator
        class_ = "RegExFunctions"
        SetPattern(vbNull)
        SetTestString(vbNull)
        SetIgnoreCase False
    End Sub

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
        If vbNull = oRE.pattern Then Err.Raise 2, funct, "RegEx pattern was not set: use SetPattern()"
        If vbNull = testString Then Err.Raise 3, funct, "RegEx test string was not set: use SetTestString()"
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
