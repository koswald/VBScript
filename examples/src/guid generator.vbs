
'script for "guid generator.hta"

Option Explicit

Dim generator, output

Sub Window_OnLoad
    Self.ResizeTo 550, 200
    document.title = "GUID Generator"
    With CreateObject("VBScripting.Includer")
        Execute .read("GUIDGenerator")
    End With
    Set generator = New GUIDGenerator
    Set output = document.getElementsByTagName("input")(0)
    Generate
End Sub

Sub Generate
    output.value = generator.generate
    output.select
End Sub

Sub Window_OnUnLoad
    Set output = Nothing
    Set generator = Nothing
End Sub       
