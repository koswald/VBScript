
Class StringFormatter

    'Property Pluralize
    'Parameters: count, noun
    'Returns a string
    'Remark: Returns a string beginning with the specified count, followed by a space, then the specified noun, and finally, unless the count is 1, an "s". If the noun has irregular pluralization, pass in a two-element array for the noun: Split("person people"). Pluralize is the default property for the class.

    Public Default Property Get Pluralize(count, noun_)
        Dim s, noun : noun = noun_ : s = count & " "
        If vbString = VarType(noun) Then
            'convert string to array
            noun = Array("", "")
            noun(0) = noun_
            noun(1) = noun_ & "s"
        End If
        If count = 1 Then
            s = s & noun(0) 'singular
        Else
            s = s & noun(1) 'plural
        End If
        Pluralize = s
    End Property

End Class
