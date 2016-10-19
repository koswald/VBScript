'
'Provides a pluralizer as the default method
'
'Usage example
'' With CreateObject("includer")
''     ExecuteGlobal(.read("StringFormatter"))
'' End With
'' Dim pluralizer : Set pluralizer = New StringFormatter
'' 
'' WScript.Echo pluralizer(3, "dog") '3 dogs
'' WScript.Echo pluralizer(0, "dog") '0 dogs
'' pluralizer.SetZeroSingular
'' WScript.Echo pluralizer(0, "dog") '0 dog
'' WScript.Echo pluralizer(1, Split("person people")) '1 person
'' WScript.Echo pluralizer(2, Split("person people")) '2 people
'' WScript.Echo pluralizer.pluralize(12, "egg") '12 eggs
'
Class StringFormatter

    Private zero, singular, plural

    Sub Class_Initialize
        singular = "singular"
        plural = "plural"
        SetZeroPlural
    End Sub

    'Property Pluralize
    'Parameters: count, noun
    'Returns a string
    'Remark: Returns a string that may or may not be pluralized, depending on the specified count. If the noun has irregular pluralization, pass in a two-element array: <code> Split("person people")</code>. Otherwise, you may pass in either a singular noun as a string, <code> red herring</code>, or else a two-element array, <code> Split("red herring | red herrings", "|")</code>. Pluralize is the default property for the class, so the property name is optional.

    Public Default Property Get Pluralize(count, noun_)
        Dim s : s = count & " "
        Dim noun : noun = noun_
        If vbString = VarType(noun) Then
            'convert string to two-element array
            noun = Array("", "")
            noun(0) = Trim(noun_)
            noun(1) = Trim(noun_) & "s"
        End If
        If count > 1 Or (count = 0 And zero = plural) Then
            s = s & Trim(noun(1)) 'plural
        Else
            s = s & Trim(noun(0)) 'singular
        End If
        Pluralize = s
    End Property

    'Method SetZeroSingular
    'Remark: Optional. Changes the default behavior of considering a count of zero to be plural.
    Sub SetZeroSingular : zero = singular : End Sub

    'Method SetZeroPlural
    'Remark: Optional. Restores the default behavior of considering a count of zero to be plural.
    Sub SetZeroPlural : zero = plural : End Sub

End Class
