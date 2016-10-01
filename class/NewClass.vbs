
Class NewClass

    Private oVBSNatives

    Private Sub Class_Initialize 'event fires on object instantiation

        With CreateObject("includer") : On Error Resume Next
            ExecuteGlobal(.read("VBSNatives"))
        End With : On Error Goto 0

        Set oVBSNatives = New VBSNatives

    End Sub

    Property Get natives : Set natives = n : End Property
    Property Get n : Set n = oVBSNatives : End Property
    Property Get sh : Set sh = n.sh : End Property
    Property Get fso : Set fso = n.fso : End Property
    Property Get a : Set a = n.a : End Property

End Class
