
'An enum and wrapper for WScript.Shell.SpecialFolders

'Usage example
'<pre>     With CreateObject( "VBScripting.Includer" ) <br />         Execute .Read( "SpecialFolders" ) <br />     End With <br />   <br />     Dim sf : Set sf = New SpecialFolders <br />     MsgBox sf.GetPath(sf.AllUsersDesktop) 'C:\Users\Public\Desktop </pre>
'
Class SpecialFolders 'use as a kind of enum with WScript.Shell.SpecialFolders

    Private sh

    Sub Class_Initialize
        Set sh = CreateObject( "WScript.Shell" )
    End Sub

    'Property GetPath
    'Parameter: a special folder alias
    'Returns a folder path
    'Remark: Returns the absolute path of the specified special folder. This is the default property, so the property name is optional.

    Public Default Property Get GetPath(alias)
        GetPath = sh.SpecialFolders(alias)
    End Property

    'Property GetAliasList
    'Returns a string
    'Remark: Returns a comma + space delimited list of the aliases of all the special folders.

    Property Get GetAliasList : GetAliasList = _
        "AllUsersDesktop, AllUsersStartMenu, AllUsersPrograms, " & _
        "AllUsersStartup, Desktop, Favorites, Fonts, MyDocuments, " & _
        "NetHood, PrintHood, Programs, Recent, SendTo, StartMenu, " & _
        "Startup, Templates"
    End Property

    'Property GetAliasArray
    'Returns an array of strings
    'Remark: Returns an array of the aliases of all the special folders.

    Property Get GetAliasArray
        Dim arr : arr = Split( GetAliasList, "," )
        Dim i : For i = 0 To UBound(arr)
            arr(i) = Trim(arr(i)) 'trim off spaces
        Next
        GetAliasArray = arr
    End Property

    'Property AllUsersDesktop
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get AllUsersDesktop : AllUsersDesktop = "AllUsersDesktop" : End Property
    'Property AllUsersStartMenu
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get AllUsersStartMenu : AllUsersStartMenu = "AllUsersStartMenu" : End Property
    'Property AllUsersPrograms
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get AllUsersPrograms : AllUsersPrograms = "AllUsersPrograms" : End Property
    'Property AllUsersStartup
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get AllUsersStartup : AllUsersStartup = "AllUsersStartup" : End Property
    'Property Desktop
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get Desktop : Desktop = "Desktop" : End Property
    'Property Favorites
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get Favorites : Favorites = "Favorites" : End Property
    'Property Fonts
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get Fonts : Fonts = "Fonts" : End Property
    'Property MyDocuments
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get MyDocuments : MyDocuments = "MyDocuments" : End Property
    'Property NetHood
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get NetHood : NetHood = "NetHood" : End Property
    'Property PrintHood
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get PrintHood : PrintHood = "PrintHood" : End Property
    'Property Programs
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get Programs : Programs = "Programs" : End Property
    'Property Recent
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get Recent : Recent = "Recent" : End Property
    'Property SendTo
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get SendTo : SendTo = "SendTo" : End Property
    'Property StartMenu
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get StartMenu : StartMenu = "StartMenu" : End Property
    'Property Startup
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get Startup : Startup = "Startup" : End Property
    'Property Templates
    'Returns a string
    'Remark: Returns a special folder alias having the exact same characters as the property name
    Property Get Templates : Templates = "Templates" : End Property

    Sub Class_Terminate
        Set sh = Nothing
    End Sub

End Class
