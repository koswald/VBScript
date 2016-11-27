
'A timer

Class VBSTimer

    Private startDate 'Date subtype
    Private start 'single; seconds
    Private precision 'integer; decimal places

    Sub Class_Initialize
        Reset
        SetPrecision 1
    End Sub

    'Function Split
    'Returns a rounded number (Single)
    'Remark: Returns the seconds elapsed since object instantiation or since calling the Reset method. Split is the default Property.

    Public Default Function Split
        Dim endDate : endDate = Now
        daysElapsed = DateDiff("d", startDate, endDate)
        Split = Round(Timer - start + daysElapsed * 24 * 60 * 60, precision)
    End Function

    'Method SetPrecision
    'Parameter: a non-negative integer
    'Remark: Sets the number of decimal places to round the Split function return value. Precision greater than 2 should be used for comparison only due to accuracy limits.

    Sub SetPrecision(newPrecision) : precision = newPrecision : End Sub

    'Method Reset
    'Remark: Sets the timer to zero.

    Sub Reset
        startDate = Now
        start = Timer  'the Timer function returns the number of seconds elapsed since midnight.
    End Sub

End Class
