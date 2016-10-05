
'Provides functions relating to the Windows&reg; registry

'Official StdRegProv docs <a href="https://msdn.microsoft.com/en-us/library/aa393664(v=vs.85).aspx">here</a>
'

Class RegistryUtility

    Private pc, oStdRegProv

    Sub Class_Initialize
        SetPC(".") 'this also initializes or reinitializes oStdRegProv
    End Sub

    Private Property Get reg : Set reg = oStdRegProv : End Property

    'Function GetStringValue
    'Parameters: rootKey, subKey, valueName
    'Returns the value of the specified registry entry.
    'Remark: The specified registry entry must be of type string (REG_SZ). <p>Set valueName to vbEmpty or "" to retrieve the default value.</p> For rootKey, use Property HKLM or HKCU.

    Function GetStringValue(rootKey, subKey, valueName)
        reg.GetStringValue rootKey, subKey, valueName, value
        GetStringValue = value
    End Function

    'Method SetStringValue
    'Parameters: rootKey, subKey, valueName, value
    'Remark: The specified registry entry must be of type string (REG_SZ). <p>Requires elevated privileges when used with HKLM.</p> Set valueName to vbEmpty or "" for setting the default value. <br /><br />For rootKey, use Property HKLM or HKCU.

    Sub SetStringValue(rootKey, subKey, valueName, value)
        reg.SetStringValue rootKey, subKey, valueName, value
    End Sub

    'Property HKLM
    'Returns &H80000002
    'Remark: Represents HKEY_LOCAL_MACHINE. For use with the rootKey parameter.

    Property Get HKLM : HKLM = &H80000002 : End Property

    'Property HKCU
    'Returns &H80000001
    'Remark: Represents HKEY_CURRENT_USER. For use with the rootKey parameter.

    Property Get HKCU : HKCU = &H80000001 : End Property

    Private Property Get GetWmiRegToken
        GetWmiRegToken = "winmgmts:\\" & pc & "\root\default:StdRegProv"
    End Property

    'Method SetPC
    'Parameter: a computer name
    'Remark: A dot (.) can be used for the local computer (default), in place of the computer name.

    Sub SetPC(newPC)
        pc = newPC
        Set oStdRegProv = GetObject(GetWmiRegToken)
    End sub

    Sub Class_Terminate
        Set oStdRegProv = Nothing
    End Sub

End Class
