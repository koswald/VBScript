'script for "guid generator.hta"

Option Explicit
Sub Generate
    output.value = generator.generate
    output.select
End Sub

Dim generator
Sub Window_OnLoad
    Self.ResizeTo 550, 200
    document.title = "GUID Generator"
    Set includer = CreateObject("VBScripting.Includer")
    Execute includer.Read("GUIDGenerator")
    Set generator = New GUIDGenerator
    Generate
    Set includer = Nothing

    Dim includer
End Sub
Sub Window_OnUnLoad
    Set generator = Nothing
End Sub       
