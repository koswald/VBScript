
'generate a strong-name key pair

With CreateObject("includer")
    Execute .read("DotNetCompiler")
End With
With New DotNetCompiler
    .SetUserInteractive True
    .SetKeyFile "%UserProfile%\MyName.snk"
    .GenerateKeyPair
End With
