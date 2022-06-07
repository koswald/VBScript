
'Provides an object whose default property, isWoW, returns a boolean indicating whether the calling script was itself called by a SysWoW64 (32-bit) .exe file. WoW64 stands for Windows 32-bit on Windows 64-bit.
'
'How it works: .exe files in %SystemRoot%\System32 and %SystemRoot%\SysWoW64 are compared by size or checksum. If the files are the same, then the calling script is assumed to be running in a 32-bit process.
'
'Usage examples
'<pre> MsgBox New WoWChecker.BySize.isWoW <br /> MsgBox New WoWChecker.isWoW <br /> With New WoWChecker : .BySize : MsgBox .isWoW : End With <br /> With New WoWChecker.BySize : MsgBox .isWoW : End With <br /> MsgBox New WoWChecker </pre>
'
Class WoWChecker
    Private sh, fso
    Private parser
    Private Pipe64Command, Pipe32Command
    Private method, ByCheckSum_, BySize_, NotSet_

    Sub Class_Initialize
        Set sh = CreateObject( "WScript.Shell" )
        Set fso = CreateObject( "Scripting.FileSystemObject" )
        ByCheckSum_ = "ByCheckSum"
        BySize_ = "BySize"
        NotSet_ = "NotSet"
        method = NotSet_
        File = "cmd.exe"
        ByCheckSum
    End Sub

    'Property OSIs64Bit
    'Returns a boolean
    'Remark: Returns a boolean that indicates whether the Windows OS is 64-bit.
    Property Get OSIs64Bit
        Dim cmd 'Windows command
        Dim phrase 'string to search for in the output
        If "Empty" = TypeName(parser) Then
            With CreateObject( "VBScripting.Includer" )
                Execute .Read( "CommandParser" )
                Set parser = New CommandParser
            End With
        End If
        cmd = "cmd /c if defined ProgramFiles(x86) (echo x64) else (echo x86)"
        phrase = "x64"
        OSIs64Bit = parser.Result( cmd, phrase )
    End Property

    'Property isWoW
    'Returns a boolean
    'Remark: Returns a boolean that indicates whether the calling script was itself called by a SysWoW64 (32-bit) .exe file. This is the class default property.
    Public Default Property Get isWoW
        Dim pipe64, pipe32
        Set pipe64 = sh.Exec(Pipe64Command)
        Set pipe32 = sh.Exec(Pipe32Command)
        Dim out64 : out64 = GetOutput(pipe64)
        Dim out32 : out32 = GetOutput(pipe32)
        Set pipe64 = Nothing
        Set pipe32 = Nothing

        'in 32-bit mode, the files in %SystemRoot%\SysWoW64 and %SystemRoot%\System32 appear to be the same
        isWoW = (out64 = out32)
    End Property

    'Function isSysWoW64
    'Returns a boolean
    'Remark: Wraps isWoW: Same as calling isWoW.
    Function isSysWoW64 : isSysWoW64 = isWoW : End Function

    'Function isSystem32
    'Returns a boolean
    'Remark: Returns the opposite of isSysWoW64
    Function isSystem32 : isSystem32 = Not isWoW : End Function

    'Function BySize
    'Returns an object self reference
    'Remark: Optional. Specifies that the .exe files will be compared by size. BySize will not distinguish between the 32- and 64-bit .exe files if they are the same size, which is unlikely but possible. ByCheckSum is therefore more reliable.
    Function BySize
        Pipe64Command = "%ComSpec% /c dir %SystemRoot%\System32\" & File
        Pipe32Command = "%ComSpec% /c dir %SystemRoot%\SysWoW64\" & File
        method = BySize_
        Set BySize = Me
    End Function

    'Function ByCheckSum
    'Returns an object self reference
    'Remark: Selected by default. Specifies that the .exe files will be compared by checksum. ByCheckSum uses CertUtil, which ships with Windows&reg; 7 through 10, and can be manually installed on older versions.
    Function ByCheckSum
        Pipe64Command = "CertUtil -hashfile %SystemRoot%\System32\" & File & " SHA1"
        Pipe32Command = "CertUtil -hashfile %SystemRoot%\SysWoW64\" & File & " SHA1"
        method = ByCheckSum_
        Set ByCheckSum = Me
    End Function

    Private file_
    'Property File
    'Returns a string
    'Remark: Optional. Sets or gets the name of the file used in comparisons. A file by this name must be found in both %SystemRoot%\System32 and %SystemRoot%\SysWoW64. The default is <code> cmd.exe</code>.
    Public Property Get File : File = file_ : End Property
    Public Property Let File(newValue)
        If Not fso.FileExists(Expand("%SystemRoot%\System32\" & newValue)) Or Not fso.FileExists(Expand("%SystemRoot%\SysWoW64\" & newValue)) Then Err.Raise 505,, "Can't find comparison file candidate " & newValue
        file_ = newValue
    End Property

    Private Property Get Expand(compressedVar) : Expand = sh.ExpandEnvironmentStrings(compressedVar) : End Property

    Private Function GetOutput(stream)
        Dim line
        GetOutput = ""
        While Not stream.StdOut.AtEndOfStream
            line = stream.StdOut.ReadLine
            If InStr(line, File) Then
                'for the BySize method, the desired info is on the same line as the name of the File
                GetOutput = line
                If ByCheckSum_ = method Then
                    'for the ByCheckSum method, it's on the following line
                    GetOutput = stream.StdOut.ReadLine
                End If
            End If
        Wend
    End Function

    Sub Class_Terminate
        Set sh = Nothing
        Set fso = Nothing
    End Sub
End Class
