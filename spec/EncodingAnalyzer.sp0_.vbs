
'Show Byte Order Mark details for selected formats

MsgBox "UTF7" & vbTab & "hex" & vbTab & "dec" & vbLf _
     & "byte0" & vbTab & "&H2b" & vbTab & &H2b & vbLf _
     & "byte1" & vbTab & "&H2f" & vbTab & &H2f & vbLf _
     & "byte2" & vbTab & "&H76" & vbTab & &H76 & vbLf & vbLf _
     _
     & "UTF8" & vbTab & "hex" & vbTab & "dec" & vbLf  _
     & "byte0" & vbTab & "&Hef" & vbTab & &Hef & vbLf _
     & "byte1" & vbTab & "&Hbb" & vbTab & &Hbb & vbLf _
     & "byte2" & vbTab & "&Hbf" & vbTab & &Hbf & vbLf & vbLf _
     _
     & "UTF32" & vbTab & "hex" & vbTab & "dec" & vbLf  _
     & "byte0" & vbTab & "&H0" & vbTab & &H0 & vbLf _
     & "byte1" & vbTab & "&H0" & vbTab & &H0 & vbLf _
     & "byte2" & vbTab & "&Hfe" & vbTab & &Hfe & vbLf _
     & "byte3" & vbTab & "&Hff" & vbTab & &Hff _
     _
     , vbInformation _
     , "Selected Byte Order Marks"
