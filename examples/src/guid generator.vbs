'script for "guid generator.hta"

Option Explicit
Dim generator 'GUIDGenerator object

Sub Window_OnLoad
    Self.ResizeTo 550, 200
    document.title = "GUID Generator"
    Set generator = New GUIDGenerator
    With New Configurer
        If "lower" = LCase( .Item( "case" )) Then
            generator.SetLowerCase
        End If
    End With
    Generate
End Sub

Sub Generate
    output.value = generator.generate
    output.select
End Sub
