'Test the VBScripting.StringFormatter Windows Script Component

Option Explicit
Dim f 'The VBScripting.StringFormatter object under test
Dim incl 'a VBScripting.Includer object
Dim errDescr 'string: an error description
Dim x 'string

Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "VBScripting.StringFormatter Windows Script Component"
        Set f = CreateObject( "VBScripting.StringFormatter" )

    .It "should pluralize a regular noun with a count of 0, by default"
        .AssertEqual f.pluralize( 0, "erring test" ), "0 erring tests"

    .It "should pluralize an irregular noun with a count of 0, by default"
        .AssertEqual f.pluralize( 0, Split("person people" )), "0 people"

    .It "should not pluralize a reg noun, count=0, after SetZeroSingular"
        f.SetZeroSingular
        .AssertEqual f.pluralize( 0, "erring test" ), "0 erring test"

    .It "should not pluralize an irreg noun, count=0, after SetZeroSingular"
        .AssertEqual f.pluralize( 0, Split("person people" )), "0 person"

    .It "should pluralize a regular noun with a count > 1"
        .AssertEqual f.pluralize( 3, "erring test" ), "3 erring tests"

    .It "should pluralize an irregular noun with a count > 1"
        .AssertEqual f.pluralize( 3, Split("person people" )), "3 people"

    .It "should not pluralize a regular noun with count of 1"
        .AssertEqual f.pluralize( 1, "test file" ), "1 test file"

    .It "should not pluralize an irregular noun with count of 1"
        .AssertEqual f.Pluralize( 1, Split("person people" )), "1 person"

    .It "should trim spaces with count of 1, regular noun"
        .AssertEqual f.pluralize( 1, " dog " ), "1 dog"

    .It "should trim spaces with count > 1, regular noun"
        .AssertEqual f.pluralize( 2, " dog " ), "2 dogs"

    .It "should trim spaces with count of 1, irregular noun"
        .AssertEqual f.pluralize( 1, Split("person | people", "|" )), "1 person"

    .It "should trim spaces with count > 1, irregular noun"
        .AssertEqual f.pluralize( 2, Split("person | people", "|" )), "2 people"

    .It "should return a formatted string"
        .AssertEqual f.format( Array("Test ""%s"" %s", "str1", "str2" )), "Test ""str1"" str2"

    .It "should have a default member, the Format property"
        .AssertEqual f( Array("Test ""%s"" %s", "str1", "str2" )), "Test ""str1"" str2"

    .It "should return a formatted string with positive integers"
        .AssertEqual f.format(Array("Test ""%s"" %s", 1, 5)), "Test ""1"" 5"

    .It "should return a formatted string with negative integers"
        .AssertEqual f.format(Array("Test ""%s"" %s", -1, -5)), "Test ""-1"" -5"

    .It "should return a formatted string with positive singles"
        .AssertEqual f.format(Array("Test ""%s"" %s", 1.45, 5.45)), "Test ""1.45"" 5.45"

    .It "should return a formatted string with negative singles"
        .AssertEqual f.format(Array("Test ""%s"" %s", -1.45, -5.45)), "Test ""-1.45"" -5.45"

    .It "should raise an error if there are too many surrogates"
        On Error Resume Next
            x = f.format(Array("Test ""%s"" %s %s", -1.45, -5.45))
            .AssertErrorRaised
            errDescr = Err.Description
        On Error Goto 0

    .It "should return the expected error description (too many)"
        .AssertEqual errDescr, "There are too many instances of %s" & vbLf & "Pattern: Test ""%s"" %s %s"

    .It "should raise an error if there are too few surrogates"
        On Error Resume Next
            x = f.format(Array("Test ""%s"" ", -1.45, -5.45))
            .AssertErrorRaised
            errDescr = Err.Description
        On Error Goto 0

    .It "should return the expected error description (too few)"
        .AssertEqual errDescr, "There are too few instances of %s" & vbLf & "Pattern: Test ""%s"" "

End With
