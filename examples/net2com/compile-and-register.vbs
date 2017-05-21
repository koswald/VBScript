
'Example showing how to compile and register a .cs file

'To illustrate, drag WSHEventLogger.cs onto this script

'For another illustration, double-click compile-and-register-vox.bat

Main
Sub Main
    With CreateObject("includer")
        Execute(.read("DotNetCompiler"))
    End With

    With New DotNetCompiler

        'initialize
        .SetUserInteractive False
        .RestartIfNotPrivileged
        .SetOnUserCancelQuitScript True

        'the following two lines illustrate hardcoding the source file and reference; for an example of how to use command-line arguments insteadm, see compile-and-register-vox.bat 
        '.SetSourceFile "Vox.cs"
        '.AddRef "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\System.Speech.dll"

        'generate a strong-name key pair
        .GenerateKeyPair

        'compile and register x64 .dll
        .SetTargetFolder "lib\" & GetPCName & "\x64"
        .Compile
        .Register

        'compile and register x86 .dll
        .SetTargetFolder "lib\" & GetPCName & "\x86"
        .SetBitness 32
        .SetTargetName .GetTargetName & "32"
        .Compile
        .Register
    End With
End Sub

'With the git repo on Google Drive, and compiling from
'two computers sharing one Google account,
'the folder name was adjusted in an unsuccessful attempt to
'keep folders from being renamed with an appended (1) or (2), etc.
'however, using this function helps to minimize confusion

Function GetPCName
    Dim net : Set net = CreateObject("WScript.Network")
    GetPCName = LCase(net.ComputerName)
    Set net = Nothing
End Function
