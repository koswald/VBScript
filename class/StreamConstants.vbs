
'StreamConstants.vbs

'Provides constants for use with Scripting.FileSystemObject.OpenTextFile

Class StreamConstants

    'Property iForReading
    'Returns 1
    'Remark: For use with OpenTextFile argument #2
    Property Get iForReading : iForReading = 1 : End Property
    'Property iForWriting
    'Returns 2
    'Remark: For use with OpenTextFile argument #2
    Property Get iForWriting: iForWriting = 2 : End Property
    'Property iForAppending
    'Returns 8
    'Remark: For use with OpenTextFile argument #2
    Property Get iForAppending : iForAppending = 8 : End Property
    'Property bCreateNew
    'Returns True
    'Remark: For use with OpenTextFile argument #3
    Property Get bCreateNew : bCreateNew = True : End Property
    'Property bDontCreateNew
    'Returns False
    'Remark: For use with OpenTextFile argument #3
    Property Get bDontCreateNew : bDontCreateNew = False : End Property
    'Property tbAscii
    'Returns 0
    'Remark: For use with OpenTextFile argument #4
    Property Get tbAscii : tbAscii = 0 : End Property
    'Property tbUnicode
    'Returns -1
    'Remark: For use with OpenTextFile argument #4
    Property Get tbUnicode : tbUnicode = -1 : End Property
    'Property tbSystemDefault
    'Returns -2
    'Remark: For use with OpenTextFile argument #4
    Property Get tbSystemDefault : tbSystemDefault = -2 : End Property

End Class

