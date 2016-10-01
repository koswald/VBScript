
Class VBSEnvironment

    'TODO sort array so that longer variables get expanded first:
    'Expected: %ProgramFiles%; Actual: %HOMEDRIVE%\Program Files

    Private oVBSNatives, oVBSLogger
    Private defaults, filtered, nonFiltered
    Private userEnv, sysEnv, proEnv, volEnv

    Sub Class_Initialize
        With CreateObject("includer") : On Error Resume Next
            ExecuteGlobal(.read("VBSNatives"))
        End With : On Error Goto 0

        Set oVBSNatives = New VBSNatives
        defaults = GetDefaults
        filtered = "filtered"
        nonFiltered = "nonFiltered"
        Set userEnv = sh.Environment("user")
        Set sysEnv = sh.Environment("system")
        Set proEnv = sh.Environment("process")
        Set volEnv = sh.Environment("volatile")
    End Sub

    Property Get log : Set log = oVBSLogger : End Property
    Property Get n : Set n = oVBSNatives : End Property
    Property Get sh : Set sh = n.sh : End Property

    'Function Expand
    'Parameter: a string
    'Returns a string with environment variables expanded
    'Remark: Expands environment variable(s); e.g. convert %UserProfile% to C:\Users\user42

    Function Expand(string_)
        Expand = sh.ExpandEnvironmentStrings(string_)
    End Function

    'Property: Collapse
    'Parameters: a string
    'Returns: a string with environment variables collapsed
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
    'Parameter: varName
    'Returns the value of the specified user environment variable
    Property Get GetUserVar(varName) : GetUserVar = userEnv(varName) : End Property

    'Method RemoveUserVar
    'Parameter: varName
    'Remark: Removes a user environment variable
    Sub RemoveUserVar(varName)
        userEnv.remove varName
    End Sub

    'Method: CreateProcessVar
    'Parameters: varName, varValue
    'Remarks: Create a (temporary) process variable
    Sub CreateProcessVar(varName, varValue)
        proEnv(varName) = varValue
    End Sub

    'Method: SetProcessVar
    'Parameters: varName, varValue
    'Remarks: Sets or creates a (temporary) process environment variable
    Sub SetProcessVar(varName, varValue) : CreateProcessVar varName, varValue : End Sub

    'Property GetProcessVar
    'Parameter: varName
    'Returns the value of the specified environment variable
    Property Get GetProcessVar(varName) : GetProcessVar = proEnv(varName) : End Property

    'Method RemoveProcessVar
    'Parameter: varName
    'Remark: Removes the specified process environment variable
    Sub RemoveProcessVar(varName)
        proEnv.remove varName
    End Sub

    Property Get GetDefaults 'variables that often come pre-installed with Windows
        GetDefaults = Array( _
            "tmp", "temp", "AllUsersProfile", _
	        "AppData", "CommonProgramFiles", "ComputerName", "ComSpec", "DFSTracingOn", _
	        "FP_No_Host_Check", "HomeDrive", "HomePath", "LocalAppData", "LogOnServer", _
	        "Number_Of_Processors", "OS", "Path", "PathExt", "Processor_Architecture", _
	        "Processor_Identifier", "Processor_Level", "Processor_Revision", "ProgramData", _
	        "ProgramFiles", "PSModulePath", "Public", "SessionName", "SystemDrive", _
	        "SystemRoot", "Trace_Format_Search_Path", "UserDomain", "UserName", _
            "UserProfile", "WinDir", "Prompt", _
	        "OnlineServices", "PCBrand", "Platform", "CommonProgramFiles\(x86\)", _
	        "CommonProgramW6432", "Processor_ArchiteW6432", "ProgramFiles\(x86\)", _
	        "ProgramW6432", "ConfigSetRoot" )
    End Property

    Function IsADefault(name)
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
        Set userEnv = Nothing
        Set sysEnv = Nothing
        Set proEnv = Nothing
        Set volEnv = Nothing
    End Sub
End Class
