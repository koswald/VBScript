Dim i : i = 254
'MsgBox Hex(i)
MsgBox "UTF7" & vbLf _
     & "&H2b: " & &H2b & vbLf _
     & "&H2f: " & &H2f & vbLf & vbLf _
     _
     & "UTF8" & vbLf _
     & "&Hef: " & &Hef & vbLf _
     & "&Hbb: " & &Hbb & vbLf & vbLf _
     _
     & "UTF32" & vbLf _
     & "&H0: " & &H0 & vbLf _
     & "&H0: " & &H0 & vbLf _
     & "&Hfe: " & &Hfe & vbLf _
     & "&Hff: " & &Hff

'MsgBox 255 = &Hff