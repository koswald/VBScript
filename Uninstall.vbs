
prompt = "Click OK to uninstall the VBScript utilities."
settings = vbOKCancel + vbDefaultButton2 + vbExclamation + vbSystemModal
caption = WScript.ScriptName

If vbOK = MsgBox(prompt, settings, caption) Then
    With CreateObject("WScript.Shell")
        .Run "wscript Setup.vbs /u"
    End With
End If
