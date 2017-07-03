
Class VBSArrays

    'Function Uniques
    'Parameter: an array
    'Returns: an array
    'Remark: Returns an array with no duplicate items, given an array that may have some.
    Function Uniques(arr)
        Dim i, s
        For i = 0 To UBound(arr)
            If 0 = InStr(s, arr(i)) Then If i Then s = s & delim & arr(i) Else s = arr(i) 'this array item is unique, so add it to the string
        Next
        Uniques = Split(s, delim) 'convert to array
    End Function

    'Function RemoveFirstElement
    'Returns: an array of strings
    'Parameter: an array of strings
    'Remark: Returns a array without the first element of the specified array.
    Property Get RemoveFirstElement(arr)
        Dim i, s : s = ""
        If UBound(arr) < 1 Then RemoveFirstElement = Split("") : Exit Property 'edge case
        For i = 1 To UBound(arr)
            If i = 1 Then s = arr(i) Else s = s & delim & arr(i)
        Next
        RemoveFirstElement = Split(s, delim) 'convert to array
    End Property

    'Function CollectionToArray
    'Returns: array of strings
    'Parameter: a collection of strings
    'Remark: Can be used to convert the WScript.Arguments object to an array, for example.
    Property Get CollectionToArray(obj)
        If vbArray = VarType(obj) Or 8204 = VarType(obj) Then 
            'parameter is already an array: don't convert
            CollectionToArray = obj
            Exit Property
        End If
        'convert the collection to an array
        Dim str, strings, i : i = 0
        For Each str In obj
            If i = 0 Then strings = str Else strings = strings & delim & str
            i = i + 1
        Next
        CollectionToArray = Split(strings, delim)
    End Property

    Private delim 'delimiter

    Sub Class_Initialize
        delim = "_-_"
    End Sub
End Class
