
'Option to uninstall the VBScript Utilities/Extensions

msg = "Click OK to uninstall the VBScript utilities."
caption = "Uninstall?"
mode = vbExclamation + vbSystemModal + vbOKCancel + vbDefaultButton2

If vbOK = MsgBox(msg, mode, caption) Then
    With CreateObject("WScript.Shell")
        .Run "wscript Setup.vbs /u"
    End With
End If
