'The StartupItems class provides a way to manage the programs that run automatically when Windows is started.
'
'Creating, updating, and deleting operations that affect all users must be performed with elevated privileges or else an error will occur. See comments for the Root property.
'
Class StartupItems

    Private reg 'StdRegProv object
    Private sh 'WScript.Shell object
    Private includer 'VBScripting.Includer object

    Sub Class_Initialize
        Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
        Set sh = CreateObject( "WScript.Shell" )
        Set includer = CreateObject( "VBScripting.Includer" )
        Root = HKCU
        Key = StandardBranch
    End Sub

    Sub Class_Terminate
        Set reg = Nothing
        Set sh = Nothing
        Set includer = Nothing
    End Sub

    'Property Items
    'Returns a collection
    'Remarks: Returns a collection of startup item objects, each object having a Name and a Value property: The Value property is the Windows command that starts the program that is identified by the Name property. For 64-bit systems, one of four possible collecttions may be returned, depending on the values of the Root and Key properties: two of the four collections are for the current user (Root = HKCU, the default) and two are for the local machine or all users (Root = HKLM). There are separate collections for 64-bit programs (Key = StandardBranch, the default) and for 32-bit programs (Key = WowBranch).
    Property Get Items
        Dim i 'integer: iterator
        Dim item_ 'NameValue object
        Dim aoo 'ArrayOfObjects object: the default property (Items) returns an array of all the objects added using the Add method.
        Dim names 'array populated by the EnumValues method
        Dim types 'array populated by the EnumValues method

        Execute includer.Read( "ArrayOfObjects" )
        Set aoo = New ArrayOfObjects
        reg.EnumValues Root, Key, names, types
        If Not IsArray(names) Then
            'return an array with no elements
            Items = aoo
            Exit Property
        End If
        Execute includer.Read( "NameValue" )
        For i = 0 To UBound(names)
            Set item_ = New NameValue
            item_.Name = names(i)
            item_.Value = Item(names(i)).Value
            aoo.Add item_
        Next
        Items = aoo
    End Property
    
    'Property Item
    'Parameter: name
    'Returns an object
    'Remarks: Returns a startup item object corresponding to the specified name. Return value depends on the values of the Root and Key properties. See comments for those properties and for the Items property.
   Property Get Item(name)
        Dim item_ 'NameValue object

        Execute includer.Read( "NameValue" )
        Set item_ = New NameValue
        item_.Name = name
        item_.Value = sh.RegRead(WSHRoot & "\" & Key & "\" & name)
        Set Item = item_
    End Property

    'Method CreateItem
    'Parameters: name, command
    'Remarks: Creates a new startup item in the registry with the specified name and command. For Root = HKLM, an error will occur if privileges are not elevated. The Root and Key properties both affect where in the registry the item will be created. For 32-bit apps on a 64-bit system, use Key = WowBranch. See comments for the Items property. 
    Sub CreateItem(name, command)
        Dim erred 'boolean
        Dim errNum 'integer
        Dim errDesc 'string
        erred = False
        If "" = Trim(name) Then
            Err.Raise 5,, "The name cannot be empty."
        End If
        On Error Resume Next
            sh.RegWrite WSHRoot & "\" & Key & "\" & name, command
            If Err Then
                erred = True
                errNum = Err.Number
                errDesc = Err.Description
            End If
        On Error Goto 0
        If erred Then
            Err.Raise 17,, "Elevated privileges are required to create or update registry values in HKEY_LOCAL_MACHINE." & vbLf & vbLf & errDesc & " ( " & Hex( errNum ) & " )."
        End If
    End Sub

    'Method UpdateItem
    'Parameters: name, command
    'Remarks: Same as the CreateItem method.
    Sub UpdateItem(name, command)
        CreateItem name, command
    End Sub

    'Method RemoveItem
    'Parameter: name
    'Remarks: Same as the DeleteItem method.
    Sub RemoveItem(name)
        DeleteItem(name)
    End Sub

    'Method DeleteItem
    'Parameter: name
    'Remarks: Deletes the startup item with the specified name. For Root = HKLM, an error will occur if privileges are not elevated. The Root and Key properties both affect where in the registry the item will be deleted from. For 32-bit apps on a 64-bit system, use Key = WowBranch. See comments for the Items property. 
    Sub DeleteItem(name)
        Dim erred 'boolean
        Dim errNum 'integer
        Dim errDesc 'string
        erred = False
        On Error Resume Next
            sh.RegDelete WSHRoot & "\" & Key & "\" & name
            If Err Then
                erred = True
                errNum = Err.Number
                errDesc = Err.Description
            End If
        On Error Goto 0
        If erred Then
            Err.Raise 17,, "Elevated privileges are required to delete registry values in HKEY_LOCAL_MACHINE." & vbLf & vbLf & errDesc & " ( " & Hex( errNum ) & " )."
        End If
    End Sub

    'Property Root
    'Parameter: an integer
    'Returns: an integer
    'Remarks: Together with the Key property, gets or sets the location in the registry where items will be read from, deleted from, or written to by the other properties and methods. The Root value can be specified by the property HKCU or HKLM. Root determines whether items apply to all users (HKLM) or to the current user only (HKCU). Creating, updating, and deleting operations that affect all users must be performed with elevated privileges or else an error will occur. 
    Public Property Let Root(newRoot)
        If HKLM = newRoot Then
            WSHRoot = "HKLM"
            root_ = newRoot
        ElseIf HKCU = newRoot Then
            WSHRoot = "HKCU"
            root_ = newRoot
        Else Err.Raise 5,, "Expected Root to be &H80000002 or &H80000001."
        End If
    End Property
    Public Property Get Root
        Root = root_
    End Property
    Private root_

    'Property HKLM
    'Returns an integer
    'Remarks: Returns <code> &H80000002</code>, an integer suitable for setting the Root property. HKLM corresponds to HKEY_LOCAL_MACHINE, the system-wide all-users registry hive.
    Property Get HKLM
        HKLM = &H80000002
    End Property

    'Property HKCU
    'Returns an integer
    'Remarks: Returns <code> &H80000001</code>, an integer suitable for setting the Root property. HKCU corresponds to HKEY_CURRENT_USER, the registry hive that contains information applicable only to the current user. <strong> Note:</strong> If the current user is not a member of the Administrators group, then the current user changes when privileges are elevated.
    Property Get HKCU
        HKCU = &H80000001
    End Property

    'Property Key
    'Parameter: a string
    'Returns: a string
    'Remarks: Together with the Root property, gets or sets the location in the registry where items will be read from, deleted from, or written to by the other properties and methods. The Key value can be specified by the property StandardBranch (the default) or WowBranch.
    Public Property Let Key(newKey)
        If LCase(StandardBranch) <> LCase(newKey) _
        And LCase(WoWBranch) <> LCase(newKey) Then
            Err.Raise 5,, "The StartupItems class requires the Key property to be equivalent to either Software\Microsoft\Windows\CurrentVersion\Run or Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run"
        End If
        key_ = newKey
    End Property
    Public Property Get Key
        Key = key_
    End Property
    Private key_

    'Property StandardBranch
    'Returns a string
    'Remarks: Returns the string "Software\Microsoft\Windows\CurrentVersion\Run", which partially describes a registry location that contains information about which programs start automatically on computer startup. WoWBranch and StandardBranch are the two strings suitable for setting the Key property.
    Property Get StandardBranch
        StandardBranch = "Software\Microsoft\Windows\CurrentVersion\Run"
    End Property

    'Property WoWBranch
    'Returns a string
    'Remarks: Returns the string "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run", which partially describes a registry location that contains information about which programs start automatically on computer startup. WoWBranch is used with 64-bit systems to store paths to 32-bit programs. StandardBranch and WoWBranch are the two strings suitable for setting the Key property.
    Property Get WoWBranch
        WoWBranch = "Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Run"
    End Property

    Private Property Let WSHRoot(newWSHRoot)
        WSHRoot_ = newWSHRoot
    End Property
    Public Property Get WSHRoot
        WSHRoot = WSHroot_
    End Property
    Private WSHRoot_

    'Method OpenTaskMgr
    'Remarks: Opens the Task Manager at the Startup page.
    Sub OpenTaskMgr
        sh.Run "taskmgr /7 /startup"
    End Sub

End Class
