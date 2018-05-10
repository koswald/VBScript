
'A working example of how validation can be accomplished.
'
Class VBSValidator

    'Property GetClassName
    'Returns the class name
    'Remark: Returns                           "VBSValidator". Useful for verifying Err.Source in a unit test.
    Property Get GetClassName : GetClassName = "VBSValidator" : End Property

    'Function IsBoolean
    'Parameter a boolean candidate
    'Returns a boolean
    'Remark: Returns True if the parameter is a boolean subtype; False if not.
    Function IsBoolean(pBool)
        If vbBoolean = VarType(pBool) Then
            IsBoolean = True
        Else
            IsBoolean = False
        End If
    End Function

    'Property EnsureBoolean
    'Parameter a boolean candidate
    'Returns: boolean
    'Remark: Raises an error if the parameter is not a boolean. Unless an error is raised, returns the same value passed to it.
    Property Get EnsureBoolean(pBool)
        If Not IsBoolean(pBool) Then
            Err.Raise 1, GetClassName, CStr(pBool) & ErrDescrBool
        End If
        EnsureBoolean = pBool
    End Property

    'Function IsInteger
    'Parameter: an integer candidate
    'Returns a boolean
    'Remark: Returns True if the parameter is an integer subtype; False if not.
    Function IsInteger(pInt)
        If vbInteger = VarType(pInt) Then
            IsInteger = True
        Else
            IsInteger = False
        End If
    End Function

    'Property EnsureInteger
    'Parameter: an integer candidate
    'Returns: integer
    'Remark: Raises an error if the parameter is not an integer. Unless an error is raised, returns the same value passed to it.
    Property Get EnsureInteger(pInt)
        If Not IsInteger(pInt) Then
            Err.Raise 2, GetClassName,, CStr(pInt) & ErrDescrInt
        End If
        EnsureInteger = pInt
    End Property

    'Property ErrDescrBool
    'Returns a string
    'Remark:                                   " is not a boolean." Useful for verifying Err.Description in a unit test.
    Property Get ErrDescrBool : ErrDescrBool = " is not a boolean." : End Property

    'Property ErrDescrInt
    'Returns a string
    'Remark:                                 " is not an integer." Useful for verifying Err.Description in a unit test.
    Property Get ErrDescrInt : ErrDescrInt = " is not an integer." : End Property

End Class
