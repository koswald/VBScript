
Class VBSArrays

    'Function Uniques
    'Parameter: an array
    'Returns an array
    'Remark: Returns an array with no duplicate items, given an array that may have some.

    Function Uniques(arr)
        Dim i, s
        For i = 0 To UBound(arr)
            If 0 = InStr(s, arr(i)) Then If i Then s = s & delim & arr(i) Else s = arr(i) 'this array item is unique, so add it to the string
        Next
        Uniques = Split(s, delim) 'convert to array
    End Function

    Private Property Get delim : delim = "_-_" : End Property 'delimiter

End Class
