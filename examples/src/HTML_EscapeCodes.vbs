'Script for HTML_EscapeCodes.hta

Option Explicit
Dim format 'VBScripting.StringFormatter object

Sub BodyOnLoad
    Set format = CreateObject( "VBScripting.StringFormatter" )
    errors.innerHTML = ""
End Sub

'click handler
Sub ShowResult
    errors.innerHTML = ""
    If 1 < Len( input.Value ) Then
        input.Select
        output.innerHTML = ""
        errors.innerHTML = "Just enter a single character, and then click the button or press Enter."
        Exit Sub
    ElseIf 0 = Len( input.Value ) Then
        input.blur
        output.innerHTML = ""
        Exit Sub
    End If
    output.innerHTML = format( Array( _
        "Asc( ""%s"" ) = %s <br>" & _
        "HTML escape code = &#38;#%s;", _
        input.Value, Asc( input.Value ), _
        Asc( input.Value ) _
    ))
    input.Value = ""
    input.focus
    errors.innerHTML = ""
End Sub

'keypress handler
Sub BodyOnKeyUp
    Dim prompt 'string: MsgBox message
    Dim link 'Asc function online help link
    Const Enter = 13
    Const Esc = 27
    Const F1 = 112
    If Enter = window.event.KeyCode Then
        ShowResult()
    ElseIf Esc = window.event.KeyCode Then
        Self.Close
    ElseIf F1 = window.event.KeyCode Then
        prompt = "Do you want to open the online docs for the VBScript function Asc?"
        link = "https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/xfw01fx4(v=vs.84)"
        If vbCancel = MsgBox( prompt, vbOKCancel, document.Title ) Then
            Exit Sub
        End If
        With CreateObject( "WScript.Shell" )
            .Run link
        End With
    End If
End Sub
