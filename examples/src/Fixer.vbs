'Script for Fixer.hta

Dim fx 'FixerHelper object

Sub Window_OnLoad
    Set fx = New FixerHelper
End Sub

Sub Window_OnUnload
    Set fx = Nothing
End Sub

Sub Document_OnKeyUp
    Const F1 = 112 'keyboard key commonly used to get help
    If F1 = window.event.keyCode Then
        sh.Run "Fixer.md"
    End If
End Sub

Sub enableHta_OnClick
    If enableHta.checked Then
        fsHta.disabled = False
    Else fsHta.disabled = True
    End If
End Sub

Class FixerHelper
    Private sh 'WScript.Shell object for registry methods
    Private reg 'StdRegProv object for registry methods
    Private regex 'RexExp object for regular expression methods
    Private fso 'Scripting.FileSystemObject
    Private HKCU, HKLM, wmiRootKey 'hive integers for StdRegProv methods
    Private HKCUShortName, HKLMShortName, rootKeyShortName 'hive strings for WScript.Shell methods
    Private SysWow64, System32 'literal strings for regex.Replace arg#2
    Private htaKey, vbsKey, wsfKey 'registry keys
    Private keyName 'string: unless debugging, an empty string used in building a default-value-path used by WScript.Shell methods.
    
    Sub Class_Initialize
        Self.ResizeTo 800, 400 'width, height
        Self.MoveTo 200, 300 'left, top
        Set sh = CreateObject( "WScript.Shell" )
        Set fso = CreateObject( "Scripting.FileSystemObject" )
        Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
        Set regex = New RegExp
        regex.IgnoreCase = True
        regex.Pattern = "\\(System32|SysWow64)\\"
        SysWow64 = "\SysWow64\" 'for regex.Replace arg#2
        System32 = "\System32\"
        htaKey = "Software\Classes\htafile\Shell\Open\Command"
        vbsKey = "Software\Classes\VBSFile\Shell\Open\Command"
        wsfKey = "Software\Classes\WSFFile\Shell\Open\Command"
        keyName = "" 'use empty string for the default value, some other value for debugging
        HKCU = &H80000001
        HKLM = &H80000002
        wmiRootKey = HKLM
        HKCUShortName = "HKCU"
        HKLMShortName = "HKLM"
        rootKeyShortName = HKLMShortName

        ShowVal htaKey, spanHta 'initialize label for radio button group
        ShowVal vbsKey, spanVbs
        ShowVal wsfKey, spanWsf
        InitializeHTML "hta"
        InitializeHTML "vbs"
        InitializeHTML "wsf"

        'if privileges are not elevated,  restart the .hta
        Dim admin : Set admin = CreateObject( "VBScripting.Admin" )
        If Not admin.PrivilegesAreElevated Then
            Dim sa : Set sa = CreateObject( "Shell.Application" )
            Dim application : Set application = document.getElementsByTagName( "application" )(0)
            sa.ShellExecute "mshta", application.CommandLine,, "runas"
            Set sa = Nothing
            Set admin = Nothing
            document.ParentWindow.Close
        End If
        Set admin = Nothing
        document.Title = fso.GetBaseName(document.location.href)
        enableHta_OnClick()
    End Sub

    'update HTML element with the registry value
    Sub ShowVal(key, span)
        span.innerHTML = GetRegValue(key)
    End Sub

    'Replace SysWow64 with System32 or vice versa, in the specified registry key,
    'and update the HTML element
    Sub ReplaceKeyString(key, replaceWith, span)
        SetRegValue key, regex.Replace(GetRegValue(key), replaceWith)
        ShowVal key, span
    End Sub

    'Write to the registry
    'Expected types: REG_SZ for htafile key and REG_EXPAND_SZ for the others,
    'but don't assume that this is so: use the same type that the system is using
    Sub SetRegValue(key_, val)
        Dim key : key = rootKeyShortName & "\" & key_ & "\" & keyName
        sh.RegWrite key, val, GetRegValueType(wmiRootKey, key_, keyName)
    End Sub

    'get a value from the registry
    Function GetRegValue(key_)
        Dim key : key = rootKeyShortName & "\" & key_ & "\" & keyName
        GetRegValue = sh.RegRead(key)
    End Function

    'Get registry key value type strings used by WScript.Shell RegWrite
    Function GetRegValueType(rootKey, subKey, valueName)
        Dim i, aNames, aTypes, iType, sType
        EnumValues rootKey, subKey, aNames, aTypes
        For i = 0 To UBound(aNames)
            If LCase(valueName) = LCase(aNames(i)) Then iType = aTypes(i)
        Next
        Select Case iType
        Case REG_SZ sType = "REG_SZ"
        Case REG_EXPAND_SZ sType = "REG_EXPAND_SZ"
        Case REG_BINARY sType = "REG_BINARY"
        Case REG_DWORD sType = "REG_DWORD"
        Case Else sType = "Type not supported by WScript.Shell.RegWrite"
        End Select
        GetRegValueType = sType
    End Function

    'Method EnumValues
    'Parameters: rootKey, subKey, aNames, aTypes
    'Remark: Wraps the StdRegProv EnumValues method. The aNames and aTypes variables, passed by reference are populated with arrays of key value name strings and type integers.
    Sub EnumValues(rootKey, subKey, aNames, aTypes)
        reg.EnumValues rootKey, subKey, aNames, aTypes

        'if null check fails, try again after writing a random value to the registry
        If VarType(aNames) <> vbNull Then Exit Sub
        Dim s : s = "928507A9-7958-4E6E-A0B1-C33A5D4D602A"
        On Error Resume Next
            reg.SetStringValue rootKey, subKey, s, s
            reg.EnumValues rootKey, subKey, aNames, aTypes
            reg.DeleteValue rootKey, subKey, s
        On Error Goto 0
    End Sub

    'Define the registry value type "constants". The Const keyword is not supported in a VBScript Class, if not in a Sub or Function or Property.
    Property Get REG_SZ : REG_SZ = 1 : End Property
    Property Get REG_EXPAND_SZ : REG_EXPAND_SZ = 2 : End Property
    Property Get REG_BINARY : REG_BINARY = 3 : End Property
    Property Get REG_DWORD : REG_DWORD = 4 : End Property
    Property Get REG_MULTI_SZ : REG_MULTI_SZ = 7 : End Property
    Property Get REG_QWORD : REG_QWORD = 11 : End Property

    'click handlers for the radio buttons
    Sub SetVbs64 : ReplaceKeyString vbsKey, System32, spanVbs : End Sub
    Sub SetVbs86 : ReplaceKeyString vbsKey, SysWow64, spanVbs : End Sub
    Sub SetWsf64 : ReplaceKeyString wsfKey, System32, spanWsf : End Sub
    Sub SetWsf86 : ReplaceKeyString wsfKey, SysWow64, spanWsf : End Sub
    Sub SetHta64 : ReplaceKeyString htaKey, System32, spanHta : End Sub
    Sub SetHta86 : ReplaceKeyString htaKey, SysWow64, spanHta : End Sub

    'initialize radio button status and fieldset legend for the specified file type
    Sub InitializeHTML(fileType)
        Dim x64element, x86element, legend, re, key, keyVal
        Set re = New RegExp
        re.IgnoreCase = True
        Set x64element = document.getElementById("rad64" & fileType)
        Set x86element = document.getElementById("rad86" & fileType)
        Execute "key = " & fileType & "Key"
        keyVal = GetRegValue(key)

        re.Pattern = "\\System32\\" 'regex pattern must escape \ with \
        If re.Test(keyVal) Then x64element.checked = "checked"
        re.Pattern = "\\SysWow64\\"
        If re.Test(keyVal) Then x86element.checked = "checked"

        Set legend = document.getElementById("lgd" & fileType)
        legend.innerHTML = rootKeyShortName & "\" & key

        Set x64element = Nothing
        Set x86element = Nothing
        Set re = Nothing
        Set legend = Nothing
    End Sub

    Sub Class_Terminate
        Set sh = Nothing
        Set fso = Nothing
        Set regex = Nothing
        Set reg = Nothing
    End Sub
    
End Class
