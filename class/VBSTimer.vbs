
'A timer
'
Class VBSTimer

    'Function Split
    'Returns a rounded number (Single)
    'Remark: Returns the seconds elapsed since object instantiation or since calling the Reset method. Split is the default Property.
    Public Default Function Split
        Split = Round(UnroundedSplit, precision)
    End Function

    Function UnroundedSplit
        Dim endDate : endDate = Now
        daysElapsed = DateDiff("d", startDate, endDate)
        UnroundedSplit = Timer - start + daysElapsed * 24 * 60 * 60
    End Function

    'Method SetPrecision
    'Parameter: 0, 1, or 2
    'Remark: Sets the number of decimal places to round the Split function return value. Default is 2.
    Sub SetPrecision(newPrecision)
        If Not IsNumeric(newPrecision) Then
            precision = 0
        ElseIf Abs(newPrecision) > 1.5 Then
            precision = 2
        ElseIf Abs(newPrecision) > 0.5 Then
            precision = 1
        Else
            precision = 0
        End If
    End Sub

    'Property GetPrecision
    'Returns 0, 1, or 2
    'Remark: Returns the current precision.
    Property Get GetPrecision : GetPrecision = precision : End Property

    'Method Reset
    'Remark: Sets the timer to zero.
    Sub Reset
        startDate = Now
        start = Timer  'the Timer function returns the number of seconds elapsed since midnight.
    End Sub

    Private startDate 'Date subtype
    Private start 'single; seconds
    Private precision 'integer; decimal places

    Sub Class_Initialize
        Reset
        SetPrecision 2
    End Sub

End Class
