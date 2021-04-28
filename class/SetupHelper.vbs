' Class SetupHelper

' Supports alternative, experimental, setup scenarios:

' 1. The original purpose was to provide custom registration of project Windows Script Component (.wsc) files and VBScript extension .dll files using HKey_Current_User instead of HKey_Local_Machine. For a brief explanation of why this approach was abandoned, see [SetupPerUser.md](../SetupPerUser.md).

' 2. Another alternate use was for experimental registration of .wsc (Windows Script Component) files when the registration failed after the Windows 10 feature edition 20H2 update on Windows 10 Home edition. The same behavior was not observed on Windows 10 Pro edition, or after the second Windows restart.

' If the calling script is not in the project root folder (recommended), then the ComponentFolder and ConfigFile properties must be set before calling the Setup method, specifying the paths or relative paths to the items. It is suggested that the working directory be set first, so that the other properties can be set with reference to that, without ambiguity. This can be done with the class CurrentDirectory property or by using the WScript.Shell CurrentDirectory property, or by other means.

'

Class SetupHelper ' original name SetupPerUser

    Public Property Get DefaultConfigFile
        DefaultConfigFile = "RegistrationData.config"
    End Property
    Public Property Get DefaultComponentFolder
        DefaultComponentFolder = "class\wsc"
    End Property
    Public Property Get DefaultDllFolder
        DefaultDllFolder = ".NET\lib"
    End Property

    Sub Setup
        Init
        CheckConfigData
        If unregistering Then
            UnregisterWscs
            UnregisterDlls
            Deleter.DeleteKey Root, _
                "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\VBScripting\"
            sh.PopUp "Finsihed unregistering.", 4, WScript.ScriptName, vbInformation + vbSystemModal
        Else
            CompileExtensions
            RegisterWscs
            RegisterDlls
            ProgramsAndFeaturesEntry
            sh.PopUp "Finsihed registering.", 4, WScript.ScriptName, vbInformation + vbSystemModal
        End If
    End Sub

    'Method Init
    'Remarks: Initialize certain properties, if they have not been already.
    Sub Init
        If IsEmpty(ComponentFolder) Then
            ComponentFolder = DefaultComponentFolder
        End If
        If IsEmpty(ConfigFile) Then
            ConfigFile = DefaultConfigFile
        End If
        If IsEmpty(DllFolder) Then
            DllFolder = DefaultDllFolder
        End If
     End Sub

     'Sub CheckConfigData
     Sub CheckConfigData
        EnsureValidRegData WscGuids, 2, 3, guidOffset, guidPattern
        EnsureValidRegData DllGuids, 2, 2, guidOffset, guidPattern

        Const guidPattern = "^[A-Fa-f\d]{8}-[A-Fa-f\d]{4}-[A-Fa-f\d]{4}-[A-Fa-f\d]{4}-[A-Fa-f\d]{12}$"
        Const guidOffset = -1
        Const progIdOffset = 0
    End Sub

    'Method: EnsureValidRegData
    'Parameters: arr, indexStart, indexStep, indexOffset, pattern
    'Remark: Ensure that the registration data to be entered into the registry is valid by raising an error when invalid data is found, which will stop the calling script, provided that the error is not supressed with an 'On Error Resume Next' statement. indexOffset: the integer to add to the current index, i, to get the array index of the partial class progid or partial interface progid.
    Sub EnsureValidRegData(arr, indexStart, indexStep, indexOffset, pattern)
        Set regex = New RegExp
        regex.IgnoreCase = True
        regex.Pattern = pattern
        For i = indexStart To UBound(arr) Step indexStep : Do
            className = arr(i + indexOffset)
            If Char2IsUpperCase(className) Then Exit Do 'skip interface
            If Not regex.Test(arr(i)) Then
                Err.Raise 1,, "Invalid registration data:" & vbLf & _
                    "array index: " & i & vbLf & _
                    "data       : " & arr(i)
            End If
        Loop While False : Next
        Dim regex, i, className
    End Sub

    'Method Char2IsUpperCase
    'Remarks: If the second char of the partial progid is upper case, then the type is an interface, in which case the validation may be ignored. In this project the interface is compiled into the same .dll as the associated class.
    Public Property Get Char2IsUpperCase(partialProgid)
        Dim s : s = Left(partialProgid, 2)
        s = Right(s, 1)
        Char2IsUpperCase = UCase(s) = s
    End Property

    Sub CompileExtensions
        sh.Run "powershell -File .NET\build\compile.ps1",, synchronous
    End Sub

    Sub UnregisterDlls
        data = DllGuids
        For i = 1 To UBound(data) Step 2
            className = data(i)
            progid = format(Array( "VBScripting.%s", className ))
            guid = format(Array( "{%s}", data(i + 1) ))

            Deleter.DeleteKey Root, format(Array( _
                "Software\Classes\%s\", progid _
            ))
            Deleter.DeleteKey Root, format(Array( _
                "Software\Classes\CLSID\%s\", guid _
            ))
            Deleter.DeleteKey Root, format(Array( _
                "Software\WOW6432Node\Classes\CLSID\%s\", guid _
            ))
            Deleter.DeleteKey Root, format(Array( _
                "Software\Classes\WOW6432Node\CLSID\%s\", guid _
            ))
        Next
        Dim data, i, className, progid, guid
    End Sub

    Sub RegisterDlls
        data = DllGuids
        For i = 1 To UBound(data) Step 2 : Do
            className = data(i)
            progid = format(Array( "VBScripting.%s", className ))
            If Char2IsUpperCase(className) Then Exit Do 'skip interface
            guid = format(Array( "{%s}", data(i + 1) ))
            s = Replace( DllFolder, "\", "/" )
            dllURL = format(Array( "file:///%s/%s.dll", s, className ))

            ' Classes\progid

            sh.RegWrite format(Array( _
                "%s\Software\Classes\%s\", rootString_, progid _
            )), progid

            sh.RegWrite format(Array( _
                "%s\Software\Classes\%s\CLSID\", rootString_, progid _
            )), guid

            ' Classes\CLSID\guid

            sh.RegWrite format(Array( _
                "%s\Software\Classes\CLSID\%s\", rootString_, guid _
            )), progid

            sh.RegWrite format(Array( _
                "%s\Software\Classes\CLSID\%s\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}\", rootString_, guid _
            )), ""

            key = format(Array("%s\Software\Classes\CLSID\%s\InprocServer32", rootString_, guid))
            sh.RegWrite format(Array("%s\", key)), "mscoree.dll"
            sh.RegWrite format(Array("%s\ThreadingModel", key)), "Both"
            sh.RegWrite format(Array("%s\Class", key)), progid
            sh.RegWrite format(Array("%s\Assembly", key)), format(Array( _
                "%s, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null", className _
            ))
            sh.RegWrite format(Array("%s\RuntimeVersion", key)), "v4.0.30319"
            sh.RegWrite format(Array("%s\CodeBase", key)), dllURL

            key = format(Array("%s\0.0.0.0", key))
            sh.RegWrite format(Array("%s\Class", key)), progid
            sh.RegWrite format(Array("%s\Assembly", key)), format(Array( _
                "%s, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null", className _
            ))
            sh.RegWrite format(Array("%s\RuntimeVersion", key)), "v4.0.30319"
            sh.RegWrite format(Array("%s\CodeBase", key)), dllURL

            sh.RegWrite format(Array( _
                "%s\Software\Classes\CLSID\%s\ProgId\", rootString_, guid _
            )), progid

            ' WOW...\Classes\CLSID\guid

            sh.RegWrite format(Array( _
                "%s\Software\WOW6432Node\Classes\CLSID\%s\", rootString_, guid _
            )), progid

            sh.RegWrite format(Array( _
                "%s\Software\WOW6432Node\Classes\CLSID\%s\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}\", rootString_, guid _
            )), ""

            key = format(Array( _
                "%s\Software\WOW6432Node\Classes\CLSID\%s\InprocServer32", rootString_, guid _
            ))
            sh.RegWrite format(Array("%s\", key)), "mscoree.dll"
            sh.RegWrite format(Array("%s\ThreadingModel", key)), "Both"
            sh.RegWrite format(Array("%s\Class", key)), progid
            sh.RegWrite format(Array("%s\Assembly", key)), format(Array( _
                "%s, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null", className _
            ))
            sh.RegWrite format(Array("%s\RuntimeVersion", key)), "v4.0.30319"
            sh.RegWrite format(Array("%s\CodeBase", key)), dllURL

            key = format(Array( "%s\0.0.0.0", key))
            sh.RegWrite format(Array("%s\Class", key)), progid
            sh.RegWrite format(Array("%s\Assembly", key)), format(Array( _
                "%s, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null", className _
            ))
            sh.RegWrite format(Array("%s\RuntimeVersion", key)), "v4.0.30319"
            sh.RegWrite format(Array("%s\CodeBase", key)), dllURL

            sh.RegWrite format(Array( _
                "%s\Software\WOW6432Node\Classes\CLSID\%s\ProgId\", rootString_, guid _
            )), progid

            ' Classes\WOW...\CLSID\guid

            sh.RegWrite format(Array( _
                "%s\Software\Classes\WOW6432Node\CLSID\%s\", rootString_, guid _
            )), progid

            sh.RegWrite format(Array( _
                "%s\Software\Classes\WOW6432Node\CLSID\%s\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}\", rootString_, guid _
            )), ""


            key = format(Array( _
                "%s\Software\Classes\WOW6432Node\CLSID\%s\InprocServer32", rootString_, guid _
            ))
            sh.RegWrite format(Array("%s\", key)), "mscoree.dll"
            sh.RegWrite format(Array("%s\ThreadingModel", key)), "Both"
            sh.RegWrite format(Array("%s\Class", key)), progid
            sh.RegWrite format(Array("%s\Assembly", key)), format(Array( _
                "%s, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null", className _
            ))
            sh.RegWrite format(Array("%s\RuntimeVersion", key)), "v4.0.30319"
            sh.RegWrite format(Array("%s\CodeBase", key)), dllURL

            key = format(Array( "%s\0.0.0.0", key))
            sh.RegWrite format(Array("%s\Class", key)), progid
            sh.RegWrite format(Array("%s\Assembly", key)), format(Array( _
                "%s, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null", className _
            ))
            sh.RegWrite format(Array("%s\RuntimeVersion", key)), "v4.0.30319"
            sh.RegWrite format(Array("%s\CodeBase", key)), dllURL

            sh.RegWrite format(Array( _
                "%s\Software\Classes\WOW6432Node\CLSID\%s\ProgId\", rootString_, guid _
            )), progid

        Loop While False : Next
        Dim data, i, className, progid, guid, key
    End Sub

    Sub UnregisterWscs
        data = WscGuids
        For i = 1 To UBound(data) Step 3
            className = data(i)
            progid = format(Array( "VBScripting.%s", className ))
            guid = format(Array( "{%s}", data(i + 1) ))

            Deleter.DeleteKey Root, format(Array( _
                "Software\Classes\%s\", progid _
            ))
            Deleter.DeleteKey Root, format(Array( _
                "Software\Classes\CLSID\%s\", guid _
            ))
            Deleter.DeleteKey Root, format(Array( _
                "Software\WOW6432Node\Classes\CLSID\%s\", guid _
            ))
            Deleter.DeleteKey Root, format(Array( _
                "Software\Classes\WOW6432Node\CLSID\%s\", guid _
            ))
        Next
        Dim data, i, className, progid, guid
    End Sub

    Sub RegisterWscs
        data = WscGuids

        For i = 1 To UBound(data) Step 3

            className = data(i)
            progid = format(Array( "VBScripting.%s", className ))
            guid = format(Array( "{%s}", data(i + 1) ))
            description = data(i + 2)
            wscURL = format(Array( _
                "file:///%s/%s.wsc", _
                Replace( ComponentFolder, "\", "/" ),  _
                className _
            ))

            ' Classes\progid

            On Error Resume Next
                sh.RegWrite format(Array( _
                    "%s\Software\Classes\%s\", rootString_, progid _
                )), description
                If InvalidRootError = Err.Number Then
                    WScript.StdOut.WriteLine "  Err         : " & Err.Description
                    WScript.StdOut.WriteLine "  Root string : " & rootString_
                    WScript.StdOut.WriteLine "  Are you attempting to modify HKEY_LOCAL_MACHINE without elevated privileges?"
                End If
                WScript.Quit
            On Error Goto 0
            sh.RegWrite format(Array( _
                "%s\Software\Classes\%s\CLSID\", rootString_, progid _
            )), guid

            ' Classes\CLSID\guid

            sh.RegWrite format(Array( _
                "%s\Software\Classes\CLSID\%s\", rootString_, guid _
            )), description
            sh.RegWrite format(Array( _
                "%s\Software\Classes\CLSID\%s\InprocServer32\", rootString_, guid _
            )), Expand("%SystemRoot%\System32\scrobj.dll")
            sh.RegWrite format(Array( _
                "%s\Software\Classes\CLSID\%s\InprocServer32\ThreadingModel", rootString_, guid _
            )), "Apartment"
            sh.RegWrite format(Array( _
                "%s\Software\Classes\CLSID\%s\ProgID\", rootString_, guid _
            )), progid
            sh.RegWrite format(Array( _
                "%s\Software\Classes\CLSID\%s\ScriptletURL\", rootString_, guid _
            )), wscURL
            sh.RegWrite format(Array( _
                "%s\Software\Classes\CLSID\%s\VersionIndependentProgId\", rootString_, guid _
            )), progid

            ' WOW...\Classes\CLSID

            sh.RegWrite format(Array( _
                "%s\Software\WOW6432Node\Classes\CLSID\%s\", rootString_, guid _
            )), description
            sh.RegWrite format(Array( _
                "%s\Software\WOW6432Node\Classes\CLSID\%s\InprocServer32\", rootString_, guid _
            )), Expand("%SystemRoot%\SysWow64\scrobj.dll")
            sh.RegWrite format(Array( _
                "%s\Software\WOW6432Node\Classes\CLSID\%s\InprocServer32\ThreadingModel", rootString_, guid _
            )), "Apartment"
            sh.RegWrite format(Array( _
                "%s\Software\WOW6432Node\Classes\CLSID\%s\ProgID\", rootString_, guid _
            )), progid
            sh.RegWrite format(Array( _
                "%s\Software\WOW6432Node\Classes\CLSID\%s\ScriptletURL\", rootString_, guid _
            )), wscURL
            sh.RegWrite format(Array( _
                "%s\Software\WOW6432Node\Classes\CLSID\%s\VersionIndependentProgID\", rootString_, guid _
            )), progid

            ' Classes\WOW...\CLSID

            sh.RegWrite format(Array( _
                "%s\Software\Classes\WOW6432Node\CLSID\%s\", rootString_, guid _
            )), description
            sh.RegWrite format(Array( _
                "%s\Software\Classes\WOW6432Node\CLSID\%s\InprocServer32\", rootString_, guid _
            )), Expand("%SystemRoot%\SysWow64\scrobj.dll")
            sh.RegWrite format(Array( _
                "%s\Software\Classes\WOW6432Node\CLSID\%s\InprocServer32\ThreadingModel", rootString_, guid _
            )), "Apartment"
            sh.RegWrite format(Array( _
                "%s\Software\Classes\WOW6432Node\CLSID\%s\ProgID\", rootString_, guid _
            )), progid
            sh.RegWrite format(Array( _
                "%s\Software\Classes\WOW6432Node\CLSID\%s\ScriptletURL\", rootString_, guid _
            )), wscURL
            sh.RegWrite format(Array( _
                "%s\Software\Classes\WOW6432Node\CLSID\%s\VersionIndependentProgID\", rootString_, guid _
            )), progid
        Next

        Dim data, i, progid, guid, description, wscURL
        Const InvalidRootError = &H80070005
    End Sub

    Function Expand(str)
        Expand = sh.ExpandEnvironmentStrings(str)
    End Function

    Public Property Get KeyExists(HKey_Root, key)
        parent = fso.GetParentFolderName(key)
        baseName = fso.GetFileName(key)
        result = stdRegProv.EnumKey(HKey_Root, parent, subkeys)
        For Each subkey In subkeys
            If LCase(subkey) = LCase(baseName) Then
                KeyExists = True
                Exit Property
            End If
        Next
        KeyExists = False
        Dim parent, baseName, result, subkeys, subkey
    End Property

    Sub ProgramsAndFeaturesEntry
        Const uninstKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\VBScripting"
        Dim InstallLocation : InstallLocation = fso.GetParentFolderName(WScript.ScriptFullName)
        stdRegProv.CreateKey Root, uninstKey
        stdRegProv.SetStringValue Root, uninstKey, "DisplayName", "VBScripting Utility Classes and Extensions (Current User)"
        stdRegProv.SetDWORDValue Root, uninstKey, "NoRemove", 0
        stdRegProv.SetStringValue Root, uninstKey, "UninstallString", format(Array("wscript ""%s\SetupPerUser.wsf"" /u", InstallLocation))
        stdRegProv.SetDWORDValue Root, uninstKey, "NoModify", 1
        stdRegProv.SetStringValue Root, uninstKey, "ModifyPath", ""
        stdRegProv.SetDWORDValue Root, uninstKey, "NoRepair", 0
        stdRegProv.SetStringValue Root, uninstKey, "RepairPath", format(Array("wscript ""%s\SetupPerUser.wsf""", InstallLocation)) '""
        stdRegProv.SetStringValue Root, uninstKey, "HelpLink", "https://github.com/koswald/VBScript"
        stdRegProv.SetStringValue Root, uninstKey, "InstallLocation", InstallLocation
        stdRegProv.SetDWORDValue Root, uninstKey, "EstimatedSize", 1500 'kilobytes
        stdRegProv.SetExpandedStringValue Root, uninstKey, "DisplayIcon", "%SystemRoot%\System32\wscript.exe,2"
        stdRegProv.SetStringValue Root, uninstKey, "Publisher", "Karl Oswald"
        stdRegProv.SetStringValue Root, uninstKey, "HelpTelephone", ""
        stdRegProv.SetStringValue Root, uninstKey, "Contact", ""
        stdRegProv.SetStringValue Root, uninstKey, "UrlInfoAbout", ""
        stdRegProv.SetStringValue Root, uninstKey, "DisplayVersion", ""
        stdRegProv.SetStringValue Root, uninstKey, "Comments", ""
        stdRegProv.SetStringValue Root, uninstKey, "Readme", InstallLocation & "\ReadMe.md"
        stdRegProv.SetStringValue Root, uninstKey, "InstallDate", "" ' [YYYYMMDD]
        stdRegProv.SetDWORDValue Root, uninstKey, "Version", 0
        stdRegProv.SetDWORDValue Root, uninstKey, "VersionMajor", 0
        stdRegProv.SetDWORDValue Root, uninstKey, "VersionMinor", 0
    End Sub

    Private configFile_
    Public Property Let ConfigFile(newValue)
        configFile_ = fso.GetAbsolutePathName(newValue)
        Execute fso.OpenTextFile(configFile_).ReadAll
    End Property
    Public Property Get ConfigFile
        ConfigFile = configFile_
    End Property

    Private componentFolder_
    Public Property Let ComponentFolder(newValue)
        componentFolder_ = fso.GetAbsolutePathName(newValue)
        Set Format = GetObject("script:" & componentFolder_ & "\StringFormatter.wsc")
        Set Deleter = GetObject("script:" & componentFolder_ & "\KeyDeleter.wsc")
    End Property
    Public Property Get ComponentFolder
        ComponentFolder = componentFolder_
    End Property

    Private dllFolder_
    Public Property Let DllFolder(newValue)
        dllFolder_ = fso.GetAbsolutePathName(newValue)
    End Property
    Public Property Get DllFolder
        DllFolder = dllFolder_
    End Property

    Private wscGuids_
    Public Property Let WscGuids(newValue)
        wscGuids_ = newValue
    End Property
    Public Property Get WscGuids
        WscGuids = wscGuids_
    End Property

    Private dllGuids_
    Public Property Let DllGuids(newValue)
        dllGuids_ = newValue
    End Property
    Public Property Get DllGuids
        DllGuids = dllGuids_
    End Property

    Private currentDirectory_
    Public Property Let CurrentDirectory(newValue)
        currentDirectory_ = fso.GetAbsolutePathName(newValue)
        sh.CurrentDirectory = currentDirectory_
    End Property
    Public Property Get CurrentDirectory
        CurrentDirectory = currentDirectory_
    End Property

    Private root_, rootString_
    Public Property Let Root(newValue)
        If Not ( HKLM = newValue ) And Not ( HKCU = newValue ) Then
            Err.Raise 1,, "Root must be either &H" & Hex(HKCU) & " or &H" & Hex(HKLM) & "."
        End If
        root_ = newValue
        If HKCU = root_ Then
            rootString_ = "HKCU"
        ElseIf HKLM = root_ Then
            rootString_ = "HKLM"
        End If
    End Property
    Public Property Get Root
        Root = root_
    End Property

    'Property HKCU
    'Returns &H80000001 (2147483649)
    'Remark: Returns a value suitable for use with the root parameter of the KeyExists property.
    Public Property Get HKCU : HKCU = &H80000001 : End Property
    'Property HKLM
    'Returns &H80000002 (2147483650)
    'Remark: Returns a value suitable for use with the root parameter of the KeyExists property.
    Public Property Get HKLM : HKLM = &H80000002 : End Property

    Private sh ' WScript.Shell object
    Private fso ' Scripting.FileSystemObject
    Private stdRegProv ' WMI StdRegProv object
    Public Format ' VBScripting.StringFormatter object; this Windows Scripting Component is not required to be registered when the SetupPerUser class is instantiated. Declared Public for testing.
    Private Deleter ' VBScripting.KeyDeleter object; this Windows Scripting Component is not required to be registered when the SetupPerUser class is instantiated.
    Private unregistering
    Private synchronous

    Sub Class_Initialize
        synchronous = True
        Set sh = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set stdRegProv = GetObject("winmgmts:\\.\root\CIMv2:StdRegProv")
        unregistering = False
        For Each arg In WScript.Arguments
            If "/u" = LCase(arg) Then unregistering = True
        Next
        Root = HKLM
        CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)

        Dim arg
    End Sub

End Class
