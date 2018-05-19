
Class VBSEnvironment

    'TODO sort array so that longer variables get collapsed first:
    'Expected: %ProgramFiles%; Actual: %HOMEDRIVE%\Program Files

    Private defaults, filtered, nonFiltered
    Private userEnv, sysEnv, proEnv, volEnv
    Private sh

    Sub Class_Initialize
        defaults = GetDefaults
        filtered = "filtered"
        nonFiltered = "nonFiltered"
        Set sh = CreateObject("WScript.Shell")
        Set userEnv = sh.Environment("user")
        Set sysEnv = sh.Environment("system")
        Set proEnv = sh.Environment("process")
        Set volEnv = sh.Environment("volatile")
    End Sub

    'Function Expand
    'Parameter: a string
    'Returns a string
    'Remark: Expands environment variable(s); e.g. convert %UserProfile% to C:\Users\user42
    Function Expand(string_)
        Expand = sh.ExpandEnvironmentStrings(string_)
    End Function

    'Property: Collapse
    'Parameters: a string
    'Returns: a string
    'Remark: Collapses a string that may contain one or more substrings that can be shortened to an environment variable.
    Function Collapse(str)
        Dim varName, s : s = str

        'collapse user and system variables first, filtering out the default
        'variables, intending to collapse the ones of the most interest/import
        s = CollapseByVarType(s, "user", filtered)
        s = CollapseByVarType(s, "system", filtered)
        s = CollapseByVarType(s, "process", filtered)
        Collapse = s
    End Function

    'Method: CreateUserVar
    'Parameters: varName, varValue
    'Remarks: Create or set a user environment variable
    Sub CreateUserVar(varName, varValue)
        userEnv(varName) = varValue
    End Sub

    'Method: SetUserVar
    'Parameters: varName, varValue
    'Remarks: Set or create a user environment variable
    Sub SetUserVar(varName, varValue) : CreateUserVar varName, varValue : End Sub

    'Property GetUserVar
    'Parameter: a variable name
    'Returns the variable value
    'Remark: Returns the value of the specified user environment variable
    Property Get GetUserVar(varName) : GetUserVar = userEnv(varName) : End Property

    'Method RemoveUserVar
    'Parameter: varName
    'Remark: Removes a user environment variable
    Sub RemoveUserVar(varName)
        userEnv.remove varName
    End Sub

    'Method: CreateProcessVar
    'Parameters: varName, varValue
    'Remarks: Create a process variable
    Sub CreateProcessVar(varName, varValue)
        proEnv(varName) = varValue
    End Sub

    'Method: SetProcessVar
    'Parameters: varName, varValue
    'Remarks: Sets or creates a process environment variable
    Sub SetProcessVar(varName, varValue) : CreateProcessVar varName, varValue : End Sub

    'Property GetProcessVar
    'Parameter: varName
    'Returns: the variable value
    'Remark: Returns the value of the specified environment variable
    Property Get GetProcessVar(varName) : GetProcessVar = proEnv(varName) : End Property

    'Method RemoveProcessVar
    'Parameter: varName
    'Remark: Removes the specified process environment variable
    Sub RemoveProcessVar(varName)
        proEnv.remove varName
    End Sub

    'Property GetDefaults
    'Returns an array
    'Remark: Returns an array of common environment variables pre-installed with some versions of Windows&reg;. Not exhaustive.
    Property Get GetDefaults 'variables that often come pre-installed with Windows
        GetDefaults = Array("tmp" _
            , "temp" _
            , "AllUsersProfile" _
            , "AppData" _
            , "CommonProgramFiles" _
            , "ProgramFiles" _
            , "ProgramFiles\(x86\)" _
            , "CommonProgramFiles\(x86\)" _
            , "CommonProgramW6432" _
            , "ProgramW6432" _
            , "ComputerName" _
            , "ComSpec" _
            , "DFSTracingOn" _
            , "FP_No_Host_Check" _
            , "HomeDrive" _
            , "HomePath" _
            , "LocalAppData" _
            , "LogOnServer" _
            , "Number_Of_Processors" _
            , "OS" _
            , "Path" _
            , "PathExt" _
            , "Processor_Architecture" _
            , "Processor_Identifier" _
            , "Processor_Level" _
            , "Processor_Revision" _
            , "ProgramData" _
            , "PSModulePath" _
            , "Public" _
            , "SessionName" _
            , "SystemDrive" _
            , "SystemRoot" _
            , "Trace_Format_Search_Path" _
            , "UserDomain" _
            , "UserName" _
            , "UserProfile" _
            , "WinDir" _
            , "Prompt" _
            , "OnlineServices" _
            , "PCBrand" _
            , "Platform" _
            , "Processor_ArchiteW6432" _
            , "ConfigSetRoot" _
        )
    End Property

    Private Function IsADefault(name)
        Dim i
        For i = 0 To UBound(defaults)
            If LCase(defaults(i)) = LCase(name) Then
                IsADefault = True
                Exit Function
            End If
        Next
        IsADefault = False
    End Function

    Private Property Get CollapseByVarType(s, varType, filter)

        Dim i, arr : arr = GetVarNameArray(sh.Environment(varType), filter)

        For i = 0 To UBound(arr)
            s = CollapseOneVar(s, arr(i))
        Next

        CollapseByVarType = s
    End Property

    Private Function GetVarNameArray(envType, filter)
        Dim name, var, list : list = ""
        For Each var in envType
            name = GetName(var)
            If filter = nonFiltered Then
                list = list & " " & name
            ElseIf Not IsADefault(name) Then
                list = list & " " & name
            End If
        Next
        GetVarNameArray = Split(Trim(list))
    End Function

    Private Function GetName(varName)
        GetName = Split(varName, "=")(0)
    End Function

    Private Function CollapseOneVar(str, varName)

        Dim s : s = str
        s = TryCollapse(s, varName)

        'if no replacements were made, try upper case

        'If s = str Then s = TryCollapse(UCase(s), varName)

        'if still no replacements were made, restore the original

        'If s = UCase(str) Then s = str

        CollapseOneVar = s
    End Function

    Private Function TryCollapse(str, varName)

        Dim unexpanded : unexpanded = "%" & varName & "%"
        Dim exp : exp = sh.ExpandEnvironmentStrings(unexpanded)
        TryCollapse = Replace(str, exp, unexpanded)

    End Function

    Sub Class_Terminate
        Set sh = Nothing
        Set userEnv = Nothing
        Set sysEnv = Nothing
        Set proEnv = Nothing
        Set volEnv = Nothing
    End Sub
End Class
