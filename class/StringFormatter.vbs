
Class StringFormatter

    'Property Pluralize
    'Parameters: count, noun
    'Returns a string
    'Remark: Returns a string beginning with the specified count, followed by a space, then the specified noun, a string, and then an "s" if and only if the count is not 1. If the noun has irregular pluralization, like person and people, pass in a two-element array for the noun: Split("person|people"). Pluralize is the default property for the class.

    Public Default Property Get Pluralize(count, noun)
        If vbString = VarType(noun) Then noun = Split(noun & "|" & noun & "s", "|")
        Dim s : s = count & " "
        If count = 1 Then s = s & noun(0) Else s = s & noun(1)
        Pluralize = s
    End Property

    Sub Class_Terminate : End Sub
End Class
