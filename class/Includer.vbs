''' Includer.vbs is the script for Includer.wsc
'
'The Includer object helps with dependency management, and can be used in a .wsf, .vbs, or .hta script.
'
'How it works: The Read property returns the contents of a .vbs class file--or any other text file.
'
'Usage example

'<pre> With CreateObject( "VBScripting.Includer" )<br />     Execute .Read( "WMIUtility.vbs" ) '.vbs may be omitted<br />     Execute .Read( "TextStreamer" )<br /> End With<br /> Dim wmi : Set wmi = New WMIUtility<br /> Dim streamer : Set streamer = New TextStreamer </pre>
'
'Relative paths may be used and are relative to the location of the class folder.
'
'Registration
'
'Although Windows Script Component (.wsc) files must be registered--unless used with GetObject("script:" & AbsolutePathToWscFile)--right clicking <code> Includer.wsc</code> and selecting Register probably <strong> will not work</strong>. Instead,
'1) Run the Setup.vbs in the project folder. Or,
'2) Run the following commands in a command window with elevated privileges. The first command applies to 64-bit systems and 32-bit systems. The second command applies only to 64-bit systems.
'
'<code>     %SystemRoot%\System32\regsvr32.exe &lt;absolute-path-to&gt;\Includer.wsc </code> <br /> <code>     %SystemRoot%\SysWow64\regsvr32.exe &lt;absolute-path-to&gt;\Includer.wsc </code>
'
'<a target="_blank" href="http://github.com/koswald/VBScript/blob/master/class/wsc/ReadMe.md#user-content-registration">Alternate registration method</a>.
'
Class Includer

    Private sVersion 'version of this file
    Private sWscID 'GUID for VBScripting.Includer
    Private ForReading, DontCreateNew 'for OpenTextFile method
    Private Ascii, Unicode, SystemDefault, StreamFormat 'for OpenTextFile
    Private sh 'WScript.Shell object
    Private fso 'Scripting.FileSystemObject object
    Private analyzer 'EncodingAnalyzer object
    Private savedCurrentDirectory 'string: folder path
    Private referencePath 'path to the 'class' folder

    Private Sub Class_Initialize
        Dim path, pathX 'the path to the project folder named "class".
        Set sh = CreateObject( "WScript.Shell" )
        Set fso = CreateObject( "Scripting.FileSystemObject" )
        Set analyzer = New EncodingAnalyzer

        sVersion = "0.0.3"
        sWscID = "{ADCEC089-30DE-11D7-86BF-00606744568C}" 'must match the classid in Includer.wsc
        ForReading = 1
        Ascii = 0
        Unicode = - 1
        SystemDefault = - 2
        DontCreateNew = False

        On Error Resume Next
            pathX = sh.RegRead("HKCR\CLSID\" & sWscID & "\ScriptletURL\") 'get path to this scriptlet from the registry
        On Error Goto 0
        path = Replace(Replace(Replace(pathX, "file:///", ""), "%20", " "), "/", "\")
        referencePath = Parent(Parent(path))
        SetFormat Ascii
    End Sub

    'Function GetObj
    'Parameter: class name
    'Returns: An object
    'Remark: Returns an object based on the VBScript class with the specified name. Requires a .wsc Windows Script Component file in \class\wsc. The object does not need to be registered, although the VBScripting.Includer (this) object must be registered. See StringFormatter.wsc for an example.
    Function GetObj(className)
        'The GetObject method doesn't require that a scriptlet
        'be registered, but it does require an absolute path.
        Set GetObj = GetObject("script:" & LibraryPath & "\wsc\" & className & ".wsc")
    End Function

    'Function LoadObject
    'Parameter: class name
    'Returns: an object
    'Remark: Experimental. Returns an object based on a class (.vbs) file located in the project's <code> class</code> folder. The parameter is the class name, which is also the base name of the class .vbs file. Classes having an Init method may need to have the WScript object or the Document object passed in, using the Init method, before calling certain procedures. See the Configurer and VBSApp classes for examples of using an Init method in this way. Experimental. Does not work well when used within a Class block.
    Function LoadObject(className)
        Dim contents 'contents of a text file
        contents = PrivateRead(className)
        Execute contents
        Execute "Set LoadObject = New " & className
    End Function

    'Function Read
    'Parameter: a file
    'Return the file contents
    'Remark: Returns the contents of the specified file, which may be expressed either as an abolute path, or as a relative path relative to the <code> class</code> folder. The file name extension may be omitted for .vbs files.
    Function Read(file)

        'Expect Ascii and Unicode file formats to be mixed together in the script library...
        'If the file format is Unicode,
        'but the StreamFormat has not been set to Unicode,
        'then temporarily set the StreamFormat to Unicode,
        'read the file, then restore the previous
        'StreamFormat setting
        If analyzer.SetFile(Resolve(file)).isUTF16LE Then
            Dim savedStreamFormat
            savedStreamFormat = StreamFormat
            SetFormatUnicode
            Read = PrivateRead(file)
            StreamFormat = savedStreamFormat
        Else
            Read = PrivateRead(file)
        End If
    End Function

    'Function ReadFrom
    'Parameters: file, path
    'Returns: file contents
    'Remark: Returns the contents of the specified file, which may be expressed either as an abolute path, or as a relative path relative to the path specified. The file name extension may be omitted for .vbs files.
    Function ReadFrom(relativePath, tempReferencePath)
        Dim savedReferencePath : savedReferencePath = referencePath
        referencePath = tempReferencePath
        ReadFrom = Read(relativePath)
        referencePath = savedReferencePath
    End Function

    'Function LibraryPath
    'Returns a folder path
    'Remark: Returns the resolved, absolute path of the <code> class</code> folder, which is the reference for relative paths passed to the Read method.
    Function LibraryPath : LibraryPath = referencePath : End Function

    Sub SetLibraryPath(newPath) : referencePath = newPath : End Sub

    Sub SetFormat(newFormat) : StreamFormat = newFormat : End Sub
    Sub SetFormatAscii : SetFormat Ascii : End Sub
    Sub SetFormatUnicode : SetFormat Unicode : End Sub
    Sub SetFormatSystemDefault : SetFormat SystemDefault : End Sub

    'Return the contents of a file
    Private Function PrivateRead(file_)
        Dim file : file = Resolve(file_)
        If Not fso.FileExists(file) Then
            file = file & ".vbs" 'add the .vbs file extension and try again
            If Not fso.FileExists(file) Then
                Err.Raise 505,, "Couldn't find file """ & file & """"
                Exit Function
            End If
        End If
        Dim stream : Set stream = fso.OpenTextFile(file, ForReading, DontCreateNew, StreamFormat)
        PrivateRead = stream.ReadAll
        stream.Close
        Set stream = Nothing
    End Function

    'Resolve a relative path ("../lib/WMI.vbs") or no path => expanded, absolute path
    Private Function Resolve(path)
        SaveCurrentDirectory
        sh.CurrentDirectory = referencePath  'set the reference folder relative paths
        Resolve = fso.GetAbsolutePathName(sh.ExpandEnvironmentStrings(path))
        RestoreCurrentDirectory
    End Function

    'Get parent folder
    Private Function Parent(str)
        Parent = fso.GetParentFolderName(str)
    End Function

    Private Sub SaveCurrentDirectory : savedCurrentDirectory = sh.CurrentDirectory : End Sub
    Private Sub RestoreCurrentDirectory : sh.CurrentDirectory = savedCurrentDirectory : End Sub

End Class