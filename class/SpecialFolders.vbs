
'A kind of enum for use with WScript.Shell.SpecialFolders

'Usage example
'
''    With CreateObject("includer")
''        ExecuteGlobal(.read("SpecialFolders"))
''        ExecuteGlobal(.read("VBSNatives"))
''    End With
'
''    Dim sp : Set sp = New SpecialFolders
''    Dim n : Set n = New VBSNatives
''    MsgBox n.shell.SpecialFolders(sp.Desktop)
'
'Here is the list: AllUsersDesktop, AllUsersStartMenu, AllUsersPrograms, AllUsersStartup, Desktop, Favorites, Fonts, MyDocuments, NetHood, PrintHood, Programs, Recent, SendTo, StartMenu, Startup, Templates
'
Class SpecialFolders 'use as a kind of enum with WScript.Shell.SpecialFolders

    'Property GetList
    'Returns a string
    'Remark: Gets a list of all the properties. Comma + space delimited.

    Property Get GetList : GetList = _
        "AllUsersDesktop, AllUsersStartMenu, AllUsersPrograms, " & _
        "AllUsersStartup, Desktop, Favorites, Fonts, MyDocuments, " & _
        "NetHood, PrintHood, Programs, Recent, SendTo, StartMenu, " & _
        "Startup, Templates"
    End Property

    'Property GetArray
    'Returns an array of strings
    'Remark: Returns an array of all the special folders

    Property Get GetArray : GetArray = Split(GetList, ", ") : End Property

    Property Get AllUsersDesktop : AllUsersDesktop = "AllUsersDesktop" : End Property
    Property Get AllUsersStartMenu : AllUsersStartMenu = "AllUsersStartMenu" : End Property
    Property Get AllUsersPrograms : AllUsersPrograms = "AllUsersPrograms" : End Property
    Property Get Desktop : Desktop = "Desktop" : End Property
    Property Get AllUsersStartup : AllUsersStartup = "AllUsersStartup" : End Property
    Property Get Favorites : Favorites = "Favorites" : End Property
    Property Get Fonts : Fonts = "Fonts" : End Property
    Property Get MyDocuments : MyDocuments = "MyDocuments" : End Property
    Property Get NetHood : NetHood = "NetHood" : End Property
    Property Get PrintHood : PrintHood = "PrintHood" : End Property
    Property Get Programs : Programs = "Programs" : End Property
    Property Get Recent : Recent = "Recent" : End Property
    Property Get SendTo : SendTo = "SendTo" : End Property
    Property Get StartMenu : StartMenu = "StartMenu" : End Property
    Property Get Startup : Startup = "Startup" : End Property
    Property Get Templates : Templates = "Templates" : End Property

End Class
