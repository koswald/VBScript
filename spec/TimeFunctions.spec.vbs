
'test the TimeFunctions class

Option Explicit

With CreateObject("includer") 'get dependencies
    Execute(.read("TimeFunctions"))
    Execute(.read("TestingFramework"))
End With

Dim tf : Set tf = New TimeFunctions

With New TestingFramework

    .describe "TimeFunctions class"

    .it "should create a two-character string"

        .AssertEqual tf.TwoDigit(8), "08"

    .it "should return a day-of-week string"

        Dim date_ : date_ = "2016-09-19"
        .AssertEqual tf.DOW(date_), "Monday"

    .it "should return an abbreviated day-of-week string"

        tf.LetDOWBeAbbreviated = True
        .AssertEqual tf.DOW(date_), "Mon"

    .it "should return a day string like 2016-09-19-Mon (DOW abbreviated)"

        .AssertEqual tf.GetFormattedDay("September 19, 2016"), "2016-09-19-Mon"

    .it "should return a day string like 1970-06-07-Sunday (DOW not abbreviated)"

        tf.LetDOWBeAbbreviated = False 'restore default
        .AssertEqual tf.GetFormattedDay("June 7, 1970"), "1970-06-07-Sunday"

    .it "should return a timestring with hh:mm:ss format (24-hour)"

        .AssertEqual tf.GetFormattedTime("2016-09-19 11:45:45 PM"), "23:45:45"

End With
