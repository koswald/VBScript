
'Provides an object whose default property, isWoW, returns a boolean indicating whether the calling script was itself called by a SysWoW64 (32-bit) .exe file.

'Usage examples
''MsgBox New WoWChecker.BySize.isWoW
''MsgBox New WoWChecker.isWoW
''With New WoWChecker : .BySize : MsgBox .isWoW : End With
''With New WoWChecker.BySize : MsgBox .isWoW : End With
''MsgBox New WoWChecker
'
Class WoWChecker
    Private sh
    Private Pipe64Command, Pipe32Command
    Private method, ByCheckSum_, BySize_, NotSet_

    Sub Class_Initialize
        Set sh = CreateObject("WScript.Shell")
        ByCheckSum_ = "ByCheckSum"
        BySize_ = "BySize"
        NotSet_ = "NotSet"
        method = NotSet_
        ByCheckSum
    End Sub

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

        'in 32-bit mode, the files in %SystemRoot%\SysWoW64 and %SystemRoot%\System32 are the same

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
        Pipe64Command = "%ComSpec% /c dir %SystemRoot%\System32\cmd.exe"
        Pipe32Command = "%ComSpec% /c dir %SystemRoot%\SysWoW64\cmd.exe"
        method = BySize_
        Set BySize = Me
    End Function

    'Function ByCheckSum
    'Returns an object self reference
    'Remark: Selected by default. Specifies that the .exe files will be compared by checksum. ByCheckSum uses CertUtil, which ships with Windows&reg; 7 through 10, and can be manually installed on older versions.

    Function ByCheckSum
        Pipe64Command = "CertUtil -hashfile %SystemRoot%\System32\cmd.exe SHA1"
        Pipe32Command = "CertUtil -hashfile %SystemRoot%\SysWoW64\cmd.exe SHA1"
        method = ByCheckSum_
        Set ByCheckSum = Me
    End Function

    Private Function GetOutput(stream)
        Dim line
        GetOutput = ""
        While Not stream.StdOut.AtEndOfStream
            line = stream.StdOut.ReadLine
            If InStr(line, "cmd.exe") Then
                'for the BySize method, the desired info is on the same line as "cmd.exe"
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
    End Sub
End Class
