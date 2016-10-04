
'Regular Expression functions - work in progress

Class RegExFunctions

    Private oRE, Match, Matches, testString, class_, v, reader

    Sub Class_Initialize
        With CreateObject("includer") : On Error Resume Next
            ExecuteGlobal(.read("VBSValidator"))
        End With : On Error Goto 0
        Set oRE = New RegExp
        Set v = New VBSValidator
        class_ = "RegExFunctions"
        SetPattern(vbNull)
        SetTestString(vbNull)
    End Sub

    '''Property re
    '''Returns a reference to a RegExp object instance

    Private Property Get re : Set re = oRE : End Property

    'Method SetPattern
    'Parameter: a regex pattern
    'Remark: Sets the pattern of the RegExp object instance

    Sub SetPattern(pPattern) : oRE.pattern = pPattern : End Sub



    Sub SetTestString(pString) : testString = pString : End Sub
    Sub SetIgnoreCase(pBool) : oRE.IgnoreCase = v.EnsureBoolean(pBool, class_ & ".SetIgnoreCase") : End Sub
    Sub SetGlobal(pBool) : oRE.Global = v.EnsureBoolean(pBool, class_ & ".SetGlobal") : End Sub

    Property Get GetSubMatches
        Dim oMatch, oMatches
        Set oMatches = re.Execute(testString)
        Set oMatch = oMatches(0)
        Set GetSubMatches = oMatch.SubMatches
    End Property

    Private Sub EnsureInitialized
        Dim funct : funct = class_ & ".EnsureInitialized"
        If vbNull = oRE.pattern Then Err.Raise 2, funct, "RegEx pattern was not set: use SetPattern()"
        If vbNull = testString Then Err.Raise 3, funct, "RegEx test string was not set: use SetTestString()"
    End Sub

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
