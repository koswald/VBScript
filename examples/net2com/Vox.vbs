'verify that the .dll is registered
'tests Vox.dll or Vox32.dll, depending on the bitness of the
'wscript.exe or cscript.exe used to open this file

With CreateObject("Vox")
    .say "I just say what I'm told to say"
End With