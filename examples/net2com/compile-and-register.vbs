
'Example showing how to compile and register a .cs file

'To illustrate, drag WSHEventLogger.cs onto this script

'For another illustration, uncomment the lines with
'.SetSourceFile and .AddRef,
'then double-click this script

Main
Sub Main
    With CreateObject("includer")
        Execute(.read("DotNetCompiler"))
    End With

    With New DotNetCompiler

        'initialize
        .SetUserInteractive False
        .RestartIfNotPrivileged
        .OnUserCancelQuitScript = True
        '.SetSourceFile "Vox.cs"
        .GenerateKeyPair
        '.AddRef "C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\System.Speech.dll"

        'compile and register x64 .dll
        .SetTargetFolder "lib\x64"
        .Compile
        .Register

        'compile and register x86 .dll
        .SetTargetFolder "lib\x86"
        .SetBitness 32
        .SetTargetName .GetTargetName & "32"
        .Compile
        .Register
    End With
End Sub