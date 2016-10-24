
'test RegExFunctions.vbs

With CreateObject("includer")
    Execute(.read("RegExFunctions"))
    Execute(.read("TestingFramework"))
    Execute(.read("VBSNatives"))
End With

Dim r : Set r = New RegExFunctions

With New TestingFramework

    .describe "RegExFunctions class"

    .it "should return a reference to the RegExp object"

        Dim pattern : pattern = "lkjkjlkkj"
        r.SetPattern pattern

        .AssertEqual r.re.pattern = pattern, True

    .it "should get the first match"

        r.SetTestString "A ring of red rocks"
        r.SetPattern "r[\w]+" 'words starting with r

        .AssertEqual r.FirstMatch, "ring"

    .it "should get submatches"

        'match the first three words that start with r
        r.SetPattern "(r[\w]+).*(r[\w]+).*(r[\w]+)"

        Dim subs : Set subs = r.GetSubMatches
        Dim sub_, s : s = ""
        For Each sub_ in subs
            s = s & " " & sub_
        Next

        .AssertEqual s, " ring red rocks"

End With
