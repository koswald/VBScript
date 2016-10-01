
'A simple, working example of how validation can be accomplished.

Class VBSValidator
    
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

    'Method EnsureBoolean
    'Parameter a boolean candidate
    'Remark: Raises an error if the parameter is not a boolean
    Sub EnsureBoolean(pBool)
        If Not IsBoolean(pBool) Then 
            Err.Raise 1, GetClassName, WScript.ScriptName & ": " & CStr(pBool) & ErrDescrBool
        End If
    End Sub

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

    'Method EnsureInteger
    'Parameter: an integer candidate
    'Remark: Raises an error if the parameter is not a boolean
    Sub EnsureInteger(pInt)
        If Not IsInteger(pInt) Then 
            Err.Raise 1, GetClassName, WScript.ScriptName & ": " & CStr(pInt) & ErrDescrInt
        End If
    End Sub

    'Property ErrDescrBool
    'Returns a string
    'Remark:                                   " is not a boolean." Useful for verifying Err.Description in a unit test.
    Property Get ErrDescrBool : ErrDescrBool = " is not a boolean." : End Property

    'Property ErrDescrInt
    'Returns a string
    'Remark:                                 " is not an integer." Useful for verifying Err.Description in a unit test.
    Property Get ErrDescrInt : ErrDescrInt = " is not an integer." : End Property
       
    'Property GetClassName
    'Returns the class name
    'Remark:                                   "VBSValidator". Useful for verifying Err.Source in a unit test.
    Property Get GetClassName : GetClassName = "VBSValidator" : End Property

End Class
