'Test the RegExFunctions object

Option Explicit
Dim r 'the RegExFunctionf object under test
Dim incl 'VBScripting.Includer object
Dim subs 'a matches collection
Dim sub_ 'an item in the matches collection
Dim s 'a string

Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "RegExFunctions class"
        Set r = incl.LoadObject( "RegExFunctions" )

    .It "should return a reference to the RegExp object"
        Dim pattern : pattern = ".*'(\w+).*"
        r.SetPattern pattern
        .AssertEqual r.re.pattern = pattern, True

    .It "should get the first match"
        r.SetTestString "A ring of red rocks"
        r.SetPattern "r[\w]+"
        .AssertEqual r.FirstMatch, "ring"

    .It "should get submatches"
        'match the first three words that start with r
        r.SetPattern "(\br[\w]+).*(\br[\w]+).*(\br[\w]+)"
        Set subs = r.GetSubMatches
        s = ""
        For Each sub_ in subs
            s = s & " " & sub_
        Next
        .AssertEqual s, " ring red rocks"

End With
