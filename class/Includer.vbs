
''' Includer.vbs is the script for Includer.wsc
'
'The Includer object helps with dependency management, and can be used in a .wsf, .vbs, or .hta script.
'
'    <h5> How it works </h5>
'
'        The Read method returns the contents of a .vbs class file--or any other text file.
'
'    <h5> Usage example </h5>
'
'' With CreateObject("VBScripting.Includer")
''     Execute .read("WMIUtility.vbs")
''     Execute .read("TextStreamer") '.vbs may be omitted
'' End With
''
'' Dim wmi : Set wmi = New WMIUtility
'' Dim streamer : Set streamer = New TextStreamer
'
'        Relative paths may be used and are relative to the location of Includer.wsc.

'
'    <h5> Registration </h5>
'
'        Although Windows Script Component (.wsc) files must be registered--unless used with GetObject("script:" & AbsolutePathToWscFile)--right clicking <code> Includer.wsc</code> and selecting Register probably <strong> will not work</strong>. Instead,
'        1) Run the Setup.vbs in the project folder. Or,
'        2) Run the following commands in a command window with elevated privileges. The first command applies to 64-bit systems and 32-bit systems. The second command applies only to 64-bit systems.
'
''       %SystemRoot%\System32\regsvr32.exe &lt;absolute-path-to&gt;\Includer.wsc
''       %SystemRoot%\SysWow64\regsvr32.exe &lt;absolute-path-to&gt;\Includer.wsc
'
''''

Option Explicit : Initialize

'Function GetObj
'Parameter: className
'Returns: An object
'Remark: Returns an object based on the VBScript class with the specified name. Requires a .wsc Windows Script Component file in \class\wsc. See StringFormatter.wsc for an example.
Function GetObj(className)
    'The GetObject method doesn't require that a scriptlet 
    'be registered, but it does require an absolute path.
    Set GetObj = GetObject("script:" & LibraryPath & "\wsc\" & className & ".wsc")
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
    Dim savedReferencePath : savedReferencePath = me.referencePath
    me.referencePath = tempReferencePath
    ReadFrom = Read(relativePath)
    me.referencePath = savedReferencePath
End Function

'Function LibraryPath
'Returns a folder path
'Remark: Returns the resolved, absolute path of the folder that contains Includer.wsc, which is the reference for relative paths passed to the Read and ReadFrom methods.
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
            Err.Raise 1,, "Couldn't find file """ & file & """"
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

Const sVersion = "0.0.0"
Const sWscID = "{ADCEC089-30DE-11D7-86BF-00606744568C}" 'must match the classid
Const ForReading = 1
Const Ascii = 0
Const Unicode = - 1
Const SystemDefault = - 2
Const CreateNew = True
Const DontCreateNew = False

Dim sh, fso, StreamFormat, analyzer
Dim savedCurrentDirectory
Private referencePath

Private Sub Initialize
    Set sh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set analyzer = New EncodingAnalyzer

    'set the path against which relative paths will be referenced, i.e. the folder named "class".
    Dim thisFile
    On Error Resume Next
        thisFile = sh.RegRead("HKCR\CLSID\" & sWscID & "\ScriptletURL\") 'get path to this scriptlet from the registry
    On Error Goto 0
    thisFile = Replace(Replace(Replace(thisFile, "file:///", ""), "%20", " "), "/", "\") 'remove superfluous string
    referencePath = Parent(Parent(thisFile))
    SetFormat Ascii
End Sub
