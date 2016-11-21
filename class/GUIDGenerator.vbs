
'Generate a unique GUID

'Usage example

'' With CreateObject("includer")
''     Execute(.read("GUIDGenerator"))
''     InputBox "",, New GUIDGenerator
'' End With
'
Class GUIDGenerator

    'Property Generate
    'Returns a GUID
    'Remark: Returns a unique GUID. Generate is the default property for the class, so the property name is optional. A sample GUID: {928507A9-7958-4E6E-A0B1-C33A5D4D602A}

    Public Default Property Get Generate
        With CreateObject("Scriptlet.TypeLib")
            Generate = Left(.guid, 38)
        End With
    End Property

End Class
