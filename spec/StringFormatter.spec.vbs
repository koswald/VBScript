With CreateObject("includer")
    Execute(.read("TestingFramework"))
    Execute(.read("StringFormatter")) 'class under test
End With

Set f = New StringFormatter

With New TestingFramework

    .describe("StringFormatter class")

    .it("should pluralize a regular noun with a count of 0, by default")

        .AssertEqual f(0, "erring test"), "0 erring tests"

    .it("should pluralize an irregular noun with a count of 0, by default")

        .AssertEqual f(0, Split("person people")), "0 people"

    .it("should not pluralize a regular noun with a count of 0, after SetZeroSingular")

        f.SetZeroSingular

        .AssertEqual f(0, "erring test"), "0 erring test"

    .it("should not pluralize an irregular noun with a count of 0, after SetZeroSingular")

        .AssertEqual f(0, Split("person people")), "0 person"

    .it("should pluralize a regular noun with a count > 1")

        .AssertEqual f(3, "erring test"), "3 erring tests"

    .it("should pluralize an irregular noun with a count > 1")

        .AssertEqual f(3, Split("person people")), "3 people"

    .it("should not pluralize a regular noun with count of 1")

        .AssertEqual f(1, "test file"), "1 test file"

    .it("should not pluralize an irregular noun with count of 1")

        .AssertEqual f.Pluralize(1, Split("person people")), "1 person"

    .it("should trim spaces with count of 1, regular noun")

        .AssertEqual f(1, " dog "), "1 dog"

    .it("should trim spaces with count > 1, regular noun")

        .AssertEqual f(2, " dog "), "2 dogs"

    .it("should trim spaces with count of 1, irregular noun")

        .AssertEqual f(1, Split("person | people", "|")), "1 person"

    .it("should trim spaces with count > 1, irregular noun")

        .AssertEqual f(2, Split("person | people", "|")), "2 people"
End With
