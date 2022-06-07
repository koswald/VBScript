'Run Setup.vbs /u
'or if called with the /s (silent) switch, then
'run Setup.vbs /u /s
With WScript.Arguments.Named
   If .Exists( "s" ) Then
       args = "/u /s" ' "silent" uninstall
   Else args = "/u"
   End If
End With
With CreateObject( "Scripting.FileSystemObject" )
   parent = .GetParentFolderName( WScript.ScriptFullName )
End With
With CreateObject( "WScript.Shell" )
   .Run "wscript """ & parent & "\Setup.vbs"" " & args
End With
