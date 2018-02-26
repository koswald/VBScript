
'Generate a unique GUID

'Usage example
'<pre>With CreateObject("VBScripting.Includer")<br />     Execute .read("GUIDGenerator")<br /> End With<br /> InputBox "",, New GUIDGenerator</pre>
'
Class GUIDGenerator

    'Property Generate
    'Returns a GUID
    'Remark: Returns a unique GUID. Generate is the default property for the class, so the property name is optional. A sample GUID: {928507A9-7958-4E6E-A0B1-C33A5D4D602A}
    Public Default Property Get Generate
        With CreateObject("Scriptlet.TypeLib")
            Generate = Caseify(Left(.guid, 38))
        End With
    End Property

    'Method SetUppercase
    'Remark: Configure the Generate property to return uppercase, the default.
    Sub SetUppercase : caseness = "UCase" : End Sub
    
    'Method SetLowercase
    'Remark: Configure the Generate property to return lowercase
    Sub SetLowercase : caseness = "LCase" : End Sub

    Dim caseness

    Sub Class_Initialize
        SetUppercase
    End Sub

    Private Function Caseify(str)
        If "UCase" = caseness Then
            Caseify = UCase(str)
        ElseIf "LCase" = caseness Then
            Caseify = LCase(str)
        Else Err.Raise 1,, "Internal error in class GUIDGenerator: Incorrectly specified case. Case must be initialized in Class_Initialize."
        End If
    End Function
End Class
