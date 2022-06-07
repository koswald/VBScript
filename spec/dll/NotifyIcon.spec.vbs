'Test the VBScripting.NotifyIcon object

Option Explicit : Initialize

With New TestingFramework

    .Describe "VBScripting.NotifyIcon object"

    .It "should add a menu item"
        On Error Resume Next
            ni.AddMenuItem "Callback test", GetRef( "CallbackTest" )
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should invoke a callback to a VBScript method"
        number1 = 0
        ni.InvokeCallbackByIndex 0
        .AssertEqual number1, 1

    .It "should show the context menu"
        On Error Resume Next
            ni.ShowContextMenu
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should have a Text property"
        On Error Resume Next
            x = ni.Text
            ni.Text = x
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should have a Visible property"
        On Error Resume Next
            x = ni.Visible
            ni.Visible = x
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should have a BalloonTipTitle property"
        On Error Resume Next
            x = ni.BalloonTipTitle
            ni.BalloonTipTitle = x
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should have a BalloonTipText property"
        On Error Resume Next
            x = ni.BalloonTipText
            ni.BalloonTipText = x
            .AssertEqual Err.Description, ""
        On Error Goto 0

    .It "should return a ToolTipIcon object"
        .AssertEqual TypeName(ni.ToolTipIcon), "ToolTipIconT"

    .It "should have a Dispose method"
        On Error Resume Next
            ni.Dispose
            .AssertEqual Err.Description, ""
        On Error Goto 0

End With

Cleanup

Sub CallbackTest
    number1 = 1
End Sub

Sub Cleanup
    On Error Resume Next
        ni.Dispose
    On Error Goto 0
    Set ni = Nothing
End Sub

Dim ni
Dim number1
Dim x

Sub Initialize
    Set ni = CreateObject( "VBScripting.NotifyIcon" )
    ni.SetIconByDllFile "%SystemRoot%\System32\msdt.exe", 0, True
    With CreateObject( "VBScripting.Includer" )
        ExecuteGlobal .Read( "TestingFramework" )
    End With
End Sub
