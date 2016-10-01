
Class TimeFunctions

    Private FirstDayOfWeek, isDOWAbbreviated, oVBSValidator, class_

    Private Sub Class_Initialize
        With CreateObject("includer")' : On Error Resume Next
            ExecuteGlobal(.read("VBSValidator")) 'get class dependencies
        End With
        Set oVBSValidator = New VBSValidator

        class_ = "TimeFunctions"
        SetFirstDOW(vbSunday)
        LetDOWBeAbbreviated = False
        If Err Then MsgBox "Err.Number: " & Err & "; .Description: " & Err.Description
    End Sub

    Property Get v : Set v = oVBSValidator : End Property

    Sub SetFirstDOW(pInt)
        v.EnsureInteger(pInt)
        FirstDayOfWeek = pInt
    End Sub

    Property Let LetDOWBeAbbreviated(pBool)
        If Not v.IsBoolean(pBool) Then
            Err.Raise 1, class_, pBool & " is not a boolean"
            isDOWAbbreviated = False
        Else
            isDOWAbbreviated = pBool
        End If
    End Property

    'Return a two-char string that may have a leading 0, given a numeric integer/string/variant of length one or two

    Function TwoDigit(num)
        If IsNumeric(num) Then
            If Len(num) = 0 Then Err.Raise 1
            If Len(num) > 2 Then Err.Raise 2
            If num < 0 Then Err.Raise 3
            If Len(num) = 1 Then TwoDigit = "0" & num Else TwoDigit = num
        Else Err.Raise 4
        End If
    End Function

    'Return a day of the week string, e.g. Monday, given a VBS date

    Function DOW(someDate)
        DOW = WeekDayName(WeekDay(someDate, FirstDayOfWeek), isDOWAbbreviated, FirstDayOfWeek)
    End Function

    'Return a formatted day string; e.g. 2016-09-15-Sat

    Property Get GetFormattedDay(date_)
        GetFormattedDay = Year(date_) & "-" & TwoDigit(Month(date_)) & "-" & TwoDigit(Day(date_)) & "-" & DOW(date_)
    End Property

    'Return a formatted 24-hr time string: e.g. 13:38:45 or 00:45:32

    Property Get GetFormattedTime(date_) 'output is very similar to the native Time function
        GetFormattedTime = TwoDigit(Hour(date_)) & ":" & TwoDigit(Minute(date_)) & ":" & TwoDigit(Second(date_))
    End Property

End Class
