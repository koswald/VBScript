With WScript.Arguments.Named
   If .Exists("s") Then
      args = "/u /s" ' "silent" uninstall
   Else args = "/u"
   End If
End With
With CreateObject("WScript.Shell")
   .Run "wscript Setup.vbs " & args
End With
