'Allows for keeping configuration data for a class or script separate from the code.
'
'Requirements
'1. The configuration files are manually created with comma-delimited key/value pairs that are read/loaded into a dictionary and accessed through the Item property.
'2. The configuration files must have the <code>configure</code> filename extension. See LoadUserConfig for the one exception.
'3. The configuration files must have the same base name as the associated class file or calling script. Two exceptions: the UserConfigFile and GlobalConfigFile do not have base names.
'4. The configuration file for a script must be located in the same folder as the script.
'5. The configuration file for a class should be in the project's <code>class</code> folder, or else in another folder that is specified by the LibraryPath property. If using another folder, then the LibraryPath property must be set before calling the LoadClassConfig method or getting the ClassConfigFile property.
'6. The configuration files can have in-line or whole-line # comments.
'7. Leading and trailing whitespace is ignored in both the key and the value.
'
' Note: Three config files GlobalConfigFile, UserConfigFile, and ScriptConfigFile, are loaded in that order on instantiation of the Configure class. The most recently loaded file takes precedence if there is a conflict, so if a different precedence is desired, then the files can be reloaded in a different order. A class configuration file is loaded by the <code> LoadClassConfig</code> method or the <code> LoadFile</code> method.
'
'Example:
'
'<pre> 'Test1.vbs (located anywhere)<br /> With CreateObject( "VBScripting.Includer" )<br />     Execute .Read( "Configurer" )<br /> End With<br /> With New Configurer<br />     If .Exists( "command1" ) Then<br />         MsgBox "command1: " & .Item( "command1" )<br />     Else MsgBox "command1 key not found."<br />     End If<br /> End With</pre>
'
'<code> # Test1.configure (located in the same folder as Test1.vbs)</code>
'<code> command1, wt powershell # requires Windows Terminal</code>
'
'<pre> 'Test2.vbs (located in the "class" folder)<br /> Class Test2<br />     Sub Class_Initialize<br />         With CreateObject( "VBScripting.Includer" )<br />             Execute .Read( "Configurer" )<br />         End With<br />         With New Configurer<br />             .LoadClassConfig me<br />             If .Exists( "command2" ) Then<br />                 MsgBox .Item( "command2" )<br />             End If<br />         End With<br />     End Sub<br /> End Class</pre>
'
'<code> # Test2.configure (also located in the "class" folder)</code>
'<code> command2, pwsh # requires PowerShell 6 or higher</code>
'
Class Configurer

    Private d 'Scripting.Dictionary object
    Private sh 'WScript.Shell object
    Private fso 'Scripting.FileSystemObject
    Private includer 'VBScripting.Includer object
    Private format 'StringFormatter object
    Private currentConfigFile 'string: a filespec
    Private missingCommaMsg, missingFileMsg
    Private scr 'filespec of the calling script or .hta
    Private parent 'parent folder of the calling script or .hta

    Sub Class_Initialize
        Dim srcX 'temp string
        Set d = CreateObject( "Scripting.Dictionary" )
        Set sh = CreateObject( "WScript.Shell")
        Set fso = CreateObject( "Scripting.FileSystemObject" )
        missingCommaMsg = "The configuration file is missing a required comma. File: "
        missingFileMsg = "Couldn't find the configuration file "
        If "HTMLDocument" = TypeName( document ) Then
            scrX = Mid( document.location.href, 9 )
            scrX = Replace( scrX, "%20", " " )
            scr = Replace( scrX, "/", "\" )
        ElseIf "Object" = TypeName( WScript ) Then
            scr = WScript.ScriptFullName
        End If
        parent = fso.GetParentFolderName( scr )
        Delimiter = "|"

        'Some scripts (such as PushPrep.hta) may need to use a limited number of class procedures when the project's Setup.vbs has not been run, that is, when the following two objects are not available. So supress the errors.
        On Error Resume Next
            Set includer = CreateObject( "VBScripting.Includer" )
            Set format = CreateObject( "VBScripting.StringFormatter" )
        On Error Goto 0

        If fso.FileExists( GlobalConfigFile ) Then
            LoadGlobalConfig
        End If
        If fso.FileExists( UserConfigFile ) Then
            LoadUserConfig
        End If
        If fso.FileExists( ScriptConfigFile ) Then
            LoadScriptConfig
        End If
    End Sub

    'Property Item
    'Parameter: a key (string)
    'Returns: a value (string)
    'Remark: Returns the value of the key/value pair for the specified key. Returns Empty if the key is not found.
    Property Get Item( key )
        If d.Exists( key ) Then
            Item = d.Item( key )
        Else Item = Empty
        End If
    End Property

    'Property Count
    'Returns an integer
    'Remark: Gets the number of key/value pairs in the Configurer dictionary.
    Property Get Count
        Count = d.Count
    End Property

    'Property Exists
    'Returns a boolean
    'Parameter: a string (key)
    'Remark: Gets whether a given key/value pair exists in the Configurer dictionary. Parameter is the key.
    Property Get Exists( key )
        Exists = d.Exists( key )
    End Property

    'Property Dictionary
    'Returns an object reference
    'Remark: Returns a reference to the Configurer object's dictionary object. Properties: CompareMode, Item, Key. Methods: Add, Exists, Items, Keys, Remove, RemoveAll. See the <a target="_blank" href="https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/x4k5wbx4(v=vs.84)"> online docs</a> for the Dictionary object.
    Property Get Dictionary
        Set Dictionary = d
    End Property
    'For testing, allow resetting the dictionary
    Property Set Dictionary( newValue )
        Set d = newValue
    End Property

    'Method: LoadFile
    'Parameter: a filespec
    'Remark: Loads the specified configuration file's key/value pairs into the object's dictionary. See Item property. See also the LoadClassConfig and LoadScriptConfig methods.
    Sub LoadFile( file )
        Dim stream 'text stream to read from a file
        currentConfigFile = file

        If Not fso.FileExists( file ) Then
            Err.Raise 51,, missingFileMsg & file
        End If
        Set stream = fso.OpenTextFile( file )
        While Not stream.AtEndOfStream
            AddItem( stream.ReadLine )
        Wend
        stream.Close
    End Sub

    'Undocumented method AddItem. Parameter is a two-element, comma-delimited string. Adds the key/value pair to the dictionary object. Or if the key exists already, updates its value.
    Private Sub AddItem( str )
        Dim a
        a = ParseLine( str )
        If IsEmpty( a ) Then
            Exit Sub
        End If
        If d.Exists( a(0) ) Then
            d.Item( a(0) ) = a(1)
        Else d.Add a(0), a(1)
        End If
    End Sub

    'Undocumented function ParseLine. Returns an array whose first element is the key and second element is the value. The parameter represents a single line from the configuration file. Returns Empty for empty lines and for lines beginning with #. Removes inline # comment.
    Function ParseLine( byVal line )
        Dim a(1) 'two element array: key/value pair
        Dim j 'comma pointer
        Dim k '# pointer
        ParseLine = Empty

        ' Ignore a whole-line comment
        If "#" = Left( Trim( line ), 1 ) Then
            Exit Function

        ' Ignore a blank line
        ElseIf 0 = Len( Trim( line )) Then
            Exit Function
        End If

        'check for missing comma
        j = InStr( line, "," )
        If 0 = j Then
            Err.Raise 51,, missingCommaMsg & ConfigFile
        End If

        'check for/remove inline comment
        k = InStr( line, "#" )
        If k Then
            line = Left( line, k - 1 )
        End If
        
        'return the key/value pair
        a(0) = Trim( Left( line, j - 1 ))
        a(1) = Trim( Mid( line, j + 1 ))
        ParseLine = a
    End Function

    'Method LoadScriptConfig
    'Returns:
    'Remarks: Loads the configuration file associated with the calling script. The configuration file's key/value pairs are added to the Configurer object's dictionary object, or if the key exists already, the value is updated.
    Sub LoadScriptConfig
        LoadFile ScriptConfigFile
    End Sub

    'Property ScriptConfigFile
    'Returns a filespec
    'Remarks: Returns the filespec of the configuration file associated with the script that is using the Configurer object, the calling script or .hta. The file doesn't have to exist.
    Property Get ScriptConfigFile
        ScriptConfigFile = _
            fso.GetParentFolderName( scr ) & "\" & _
            fso.GetBaseName( scr ) & ".configure"
    End Property

    'Method LoadClassConfig
    'Parameter: a string or an object reference
    'Remarks: Loads the configuration file associated with a class file. The configuration file's key/value pairs are added to the Configurer object's dictionary object, or if the key exists already, the value is updated. The parameter may be 1) the class name, or 2) an object reference to an instance of the class, or 3) the keyword me, if called from within the class.
    Public Sub LoadClassConfig( classInfo )
        LoadFile ClassConfigFile( classInfo )
    End Sub

    'Property ClassConfigFile
    'Returns a filespec
    'Parameter: a string or an object reference.
    'Remarks: Returns the filespec of the configuration file associated with a class (.vbs) file. The file doesn't have to exist. The parameter may be 1) the class name, or 2) an object reference to an instance of the class, or 3) the keyword me, if called from within the class.
    Property Get ClassConfigFile( classInfo )
        Dim baseName
        If vbString = VarType( classInfo ) Then
            baseName = classInfo
        Else baseName = TypeName( classInfo )
        End If
        ClassConfigFile = format( Array( _
            "%s\%s.configure", _
            LibraryPath, baseName _
        ))
    End Property

    'Method LoadUserConfig
    'Remarks: Loads the user configuration file at <code>%UserProfile%&#92;.VBScripting</code>. See Note for UserConfigFile.
    Sub LoadUserConfig
        LoadFile UserConfigFile
    End Sub

    'Property UserConfigFile
    'Returns a filespec
    'Remark: Returns the filespec of a user-specific configuration file, related to the project but outside of the project folders, at <code>%UserProfile%&#92;.VBScripting</code>. The file doesn't have to exist. Note: Care should be taken when privileges are elevated and the user is not a member of the Administrators group, because as privileges are elevated, %UserProfile% changes.
    Property Get UserConfigFile
        Dim file : file = "%UserProfile%\.VBScripting"
        UserConfigFile = Expand( file )
    End Property

    Property Get Expand( str )
        Expand = sh.ExpandEnvironmentStrings( str )
    End Property

    'For testing, allow setting a new process environment variable
    Property Let MockVar( newName, newValue )
        sh.Environment( "process" )( newName ) = newValue
    End Property
    Property Get MockVar( name )
        MockVar = sh.Environment( "process" )( name )
    End Property

    'Method LoadGlobalConfig
    'Remarks: Loads the configuration file in the project folder. See comments for the GlobalConfigFile property. Equivalent to calling <code>LoadFile GlobalConfigFile</code>.
    Sub LoadGlobalConfig
        LoadFile GlobalConfigFile
    End Sub

    'Property GlobalConfigFile
    'Returns: a filespec
    'Remark: Returns the filespec of the global configuration file. The word global refers to the project only. Depending on the location of the project, the configuration file may or may not be accessible to all users. The file does not have to exist. Expected value: <code>&lt;project folder&gt;&#92;.configure</code>.
    Property Get GlobalConfigFile
        If Not IsEmpty( globalConfigFile_ ) Then
            GlobalConfigFile = globalConfigFile_
            Exit Property
        End If
        globalConfigFile_ = format( Array( _
            "%s\.configure", _
            fso.GetParentFolderName( LibraryPath ) _
        ))
        GlobalConfigFile = globalConfigFile_
    End Property
    Property Let GlobalConfigFile( newValue )
        globalConfigFile_ = newValue
    End Property
    Private globalConfigFile_

    'Property LibraryPath
    'Returns: a path
    'Remark: Gets or sets the location, i.e. the parent folder, of the class file and/or its associated configuration file. See the LoadClassConfig and LoadFile methods. Obscure. For an example, see the integration test Configurer.spec.wsf.
    Property Let LibraryPath( newPath )
        includer.SetLibraryPath newPath
    End Property
    Property Get LibraryPath
        LibraryPath = includer.LibraryPath
    End Property

    'Undocumented property. Returns the filespec of the previously loaded configuration file. Returns Empty if no file has been loaded.
    Property Get ConfigFile
        ConfigFile = currentConfigFile
    End Property

    'Trim whitespace from left and right ends of each array elment
    Sub TrimElements( byRef arr )
        Dim i
        For i = 0 To UBound( arr )
            arr( i ) = Trim( arr( i ))
        Next
    End Sub

    'Property ToArray
    'Returns an array
    'Parameter: a string
    'Remarks: Converts a string to an array. Uses the delimiter set by the Delimiter property, a vertical bar ( &#124; ) by default. Excess spaces on the left and right of each element are trimmed off.
    Property Get ToArray( str )
        array_ = Split( str, Delimiter )
        TrimElements array_
        ToArray = array_
    End Property
    Private array_

    'Property PowerShell
    'Returns a string
    'Remarks: Returns a string useful for starting a PowerShell process. If PowerShell 6 or 7 is installed, then the return value is the expanded filespec of the first "pwsh candidates" executable found that is listed in the file <code>.configure</code> in the project's root folder. If the cross-platform PowerShell is not found, returns the string <code>powershell</code>, which may be used to start a Windows PowerShell process. Since the return value may contain spaces, the string may need to be surrounded by quotes, depending on how it is used. For example, if the return value is used as the first argument of the Shell.Appliction object's ShellExecute method, then quotes are not recommended. But if the return value is used in the first argument of the WScript.Shell object's Run method, then quotes are recommended.
    Property Get PowerShell
        Dim candidate 'untested pwsh.exe filespec
        If Not IsEmpty( powershell_ ) Then
            PowerShell = powershell_
            Exit Property
        End If
        LoadGlobalConfig
        For Each candidate In ToArray( Item( "pwsh candidates" ))
            If fso.FileExists( Expand( candidate )) Then
                powershell_ = Expand( candidate )
                PowerShell = powershell_
                Exit Property
            End If
        Next
        powershell_ = PsFallback
        PowerShell = powershell_
    End Property
    Property Let PowerShell( newValue )
        powershell_ = newValue
    End Property
    Private powershell_

    'Property WT
    'Returns a string
    'Remarks: Returns the filespec of a Windows Terminal executable, if installed and listed in <code>.configure</code> in the project folder. Returns <code>Empty</code> if Windows Terminal is not installed or not found.
    Property Get WT
        Dim candidate 'string
        If Not IsEmpty( wt_ ) Then
            WT = wt_
            Exit Property
        End If
        LoadGlobalConfig
        For Each candidate In ToArray( Item( "wt candidates" ))
            If fso.FileExists( Expand( candidate )) Then
                wt_ = Expand( candidate )
                WT = wt_
                Exit Property
            End If
        Next
    End Property
    Property Let WT( newValue )
        wt_ = newValue
    End Property
    Private wt_

    'Property Delimiter
    'Returns a character
    'Remarks: Gets or sets the delimiter used in converting strings to arrays. Default is a vertical bar ( &#124; ).
    Property Let Delimiter( newValue )
        delimiter_ = newValue
    End Property
    Property Get Delimiter
        Delimiter = delimiter_
    End Property
    Private delimiter_

    'Property PsFallback
    'Returns a string
    'Remarks: Returns a ten-character string suitable for starting a Windows PowerShell process: <code>powershell</code>. This becomes the default PowerShell when the newer cross-platform PowerShell is not installed or not found.
    Property Get PsFallback
        PsFallback = "powershell"
    End Property

    'Property Init
    'Parameter: an object
    'Returns an object self-reference
    'Remarks: Initializes the Configurer object so that it can find the name of the calling script. The parameter is the WScript object, for .vbs or .wsf files, or the 'Document' object for .hta files. Required if the Configurer object was instantiated with the <a href="#includer"> VBScripting.Includer</a> object's experimental LoadObject method. Example: <pre> With CreateObject( "VBScripting.Includer" )<br />     Set c = .LoadObject( "Configurer" ).Init( WScript )<br /> End With</pre>
    Function Init( obj )
        Dim srcX 'temp string 
        If "HTMLDocument" = TypeName( obj ) Then
            scrX = Mid( document.location.href, 9 )
            scrX = Replace( scrX, "%20", " " )
            scr = Replace( scrX, "/", "\" )
        ElseIf "Object" = TypeName( obj ) Then
            scr = obj.ScriptFullName
        End If
        If fso.FileExists( ScriptConfigFile ) Then
            LoadScriptConfig
        End If
        Set Init = me
    End Function

End Class
