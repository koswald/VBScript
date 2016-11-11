
'Provides functions relating to the Windows&reg; registry
'
'Usage example
'
'' With CreateObject("includer")
''     Execute(.read("RegistryUtility"))
'' End With
'' Dim reg : Set reg = New RegistryUtility
'' Dim key : key = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
'' MsgBox reg.GetStringValue(reg.HKLM, key, "ProductName")
'
'Set valueName to vbEmpty or "" (two double quotes) to specify a key's default value.
'
'StdRegProv docs <a href="https://msdn.microsoft.com/en-us/library/aa393664(v=vs.85).aspx"> online</a>.
'
Class RegistryUtility

    Private pc, oStdRegProv

    Sub Class_Initialize
        SetPC(".")
    End Sub

    'Method SetPC
    'Parameter: a computer name
    'Remark: Optional. A dot (.) can be used for the local computer (default), in place of the computer name.

    Sub SetPC(newPC)
        pc = newPC
        Set oStdRegProv = GetObject(GetWmiRegToken)
    End sub

    Private Property Get reg : Set reg = oStdRegProv : End Property

    'Function GetStringValue
    'Parameters: rootKey, subKey, valueName
    'Returns a string
    'Remark: Returns the value of the specified registry location. The specified registry entry must be of type string (REG_SZ).

    Function GetStringValue(rootKey, subKey, valueName)
        Dim value
        reg.GetStringValue rootKey, subKey, valueName, value
        GetStringValue = value
    End Function

    'Method SetStringValue
    'Parameters: rootKey, subKey, valueName, value
    'Remark: Writes the specified REG_SZ value to the specified registry location. Writing to HKLM or HKCR requires elevated privileges.

    Sub SetStringValue(rootKey, subKey, valueName, value)
        reg.SetStringValue rootKey, subKey, valueName, value
    End Sub

    'Function GetExpandedStringValue
    'Parameters: rootKey, subKey, valueName
    'Returns a string
    'Remark: Returns the value of the specified registry location. The specified registry entry must be of type REG_EXPAND_SZ.

    Function GetExpandedStringValue(rootKey, subKey, valueName)
        Dim value
        reg.GetExpandedStringValue rootKey, subKey, valueName, value
        GetExpandedStringValue = value
    End Function

    'Method SetExpandedStringValue
    'Parameters: rootKey, subKey, valueName, value
    'Remark: Writes the specified REG_EXPAND_SZ value to the specified registry location. Writing to HKLM or HKCR requires elevated privileges.

    Sub SetExpandedStringValue(rootKey, subKey, valueName, value)
        reg.SetExpandedStringValue rootKey, subKey, valueName, value
    End Sub

    'Property HKLM
    'Returns &H80000002
    'Remark: Represents HKEY_LOCAL_MACHINE. For use with the rootKey parameter.

    Property Get HKLM : HKLM = &H80000002 : End Property

    'Property HKCU
    'Returns &H80000001
    'Remark: Represents HKEY_CURRENT_USER. For use with the rootKey parameter.

    Property Get HKCU : HKCU = &H80000001 : End Property

    'Property HKCR
    'Returns &H80000000
    'Remark: Represents HKEY_CLASSES_ROOT. For use with the rootKey parameter.

    Property Get HKCR : HKCR = &H80000000 : End Property

    Private Property Get GetWmiRegToken
        GetWmiRegToken = "winmgmts:\\" & pc & "\root\default:StdRegProv"
    End Property

    'Property GetPC
    'Returns a string
    'Remark: Returns the name of the current computer. <strong> .</strong> (dot) indicates the local computer.

    Property Get GetPC : GetPC = pc : End Property

    'Function GetRegValueType
    'Parameters: rootKey, subKey, valueName
    'Returns an integer
    'Remark: Returns a registry key value type integer.

    Function GetRegValueType(rootKey, subKey, valueName)
        Dim i, aNames, aTypes, iType, sType
        reg.EnumValues rootKey, subKey, aNames, aTypes
        For i = 0 To UBound(aNames)
            If valueName = aNames(i) Then iType = aTypes(i)
        Next
        GetRegValueType = iType
    End Function

    'Property REG_SZ
    'Returns 1
    'Remark: Returns a registry value type constant.

    Property Get REG_SZ : REG_SZ = 1 : End Property

    'Property REG_EXPAND_SZ
    'Returns 2
    'Remark: Returns a registry value type constant.

    Property Get REG_EXPAND_SZ : REG_EXPAND_SZ = 2 : End Property

    'Property REG_BINARY
    'Returns 3
    'Remark: Returns a registry value type constant.

    Property Get REG_BINARY : REG_BINARY = 3 : End Property

    'Property REG_DWORD
    'Returns 4
    'Remark: Returns a registry value type constant.

    Property Get REG_DWORD : REG_DWORD = 4 : End Property

    'Property REG_MULTI_SZ
    'Returns 7
    'Remark: Returns a registry value type constant.

    Property Get REG_MULTI_SZ : REG_MULTI_SZ = 7 : End Property

    'Property REG_QWORD
    'Returns 11
    'Remark: Returns a registry value type constant.

    Property Get REG_QWORD : REG_QWORD = 11 : End Property

    'Function GetRegValueTypeString
    'Parameters: rootKey, subKey, valueName
    'Returns a string
    'Remark: Returns a registry key value type string suitable for use with WScript.Shell RegWrite method argument #3. That is, one of "REG_SZ", "REG_EXPAND_SZ", "REG_BINARY", or "REG_DWORD".

    Function GetRegValueTypeString(rootKey, subKey, valueName)
        Select Case GetRegValueType(rootKey, subKey, valueName)
        Case REG_SZ sType = "REG_SZ"
        Case REG_EXPAND_SZ sType = "REG_EXPAND_SZ"
        Case REG_BINARY sType = "REG_BINARY"
        Case REG_DWORD sType = "REG_DWORD"
        Case Else sType = "Type not supported by WScript.Shell.RegWrite"
        End Select
        GetRegValueTypeString = sType
    End Function

    Sub Class_Terminate
        Set oStdRegProv = Nothing
    End Sub

End Class
