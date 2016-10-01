
Class StreamConstants 'for use with Scripting.FileSystemObject.OpenTextFile

    Property Get iForReading : iForReading = 1 : End Property 'n.sh.OpenTextFile arg #2
    Property Get iForWriting: iForWriting = 2 : End Property
    Property Get iForAppending : iForAppending = 8 : End Property
    Property Get bCreateNew : bCreateNew = True : End Property 'arg #3
    Property Get bDontCreateNew : bDontCreateNew = False : End Property
    Property Get tbAscii : tbAscii = 0 : End Property 'arg #4; tb => tristate boolean
    Property Get tbUnicode : tbUnicode = -1 : End Property
    Property Get tbSystemDefault : tbSystemDefault = -2 : End Property

End Class

