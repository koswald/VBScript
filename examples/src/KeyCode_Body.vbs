'Script for KeyCode.hta, body element

Sub Document_OnKeyDown
    ShowKeyCode
End Sub
Sub Document_OnKeyUp
    ShowKeyCode
End Sub
Sub ShowKeyCode
    output.innerHTML = _
        "window.event.keyCode: " & window.event.keyCode &  "<br>" & _
        "window.event.shiftKey: "& window.event.shiftKey & "<br>" & _
        "window.event.ctrlKey: " & window.event.ctrlKey &  "<br>" & _
        "window.event.altKey: "  & window.event.altKey
End Sub

