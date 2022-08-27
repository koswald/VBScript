# PopUp bug

[Overview](#overview)  
[Description](#description)  
[Steps to reproduce](#steps-to-reproduce)  
[Workaround](#workaround)

## Overview

**The bug described below is no longer present** as of project [version 1.4.0](../../ChangeLog.md#user-content-version-140). This can be verified by running [PopUp.spec.sk.vbs](./PopUp.spec.sk.vbs): The specs that begin with `exhibits` fail when the bug is not present.

## Description

Using the `With <object> ... End With` syntax in a script that also uses the WshShell.Popup method may cause the Popup return value to always be -1, even when a dialog button is pressed. The Popup method should return -1 only when the dialog times out.

## Steps to reproduce

- Create a new .vbs file with the following code.
    ```vb
    With CreateObject( "VBScripting.Includer" )
        Execute .Read( "VBSLogger" )
        Set logger = New VBSLogger
    End With
    Set sh = CreateObject( "WScript.Shell" )
    response = sh.PopUp("message", 60, "caption", vbOKCancel) '1st dialog
    MsgBox response '2nd dialog
    Set sh = Nothing
    ```
- Double click the .vbs file to run it.
- Click Cancel to the first dialog before it times out.
- If the second dialog shows -1, then the bug is present, because clicking Cancel at a Popup dialog should return 2.

> **Note**: Similar behavior is seen for other classes in the 'class' folder, and for all other buttons besides Cancel: OK, Yes, No, Abort, Retry, Ignore.

## Workaround

- Use syntax syntax similar to the following. That is, don't use the `With <object> ... End With` syntax for project classes.
    ```vb
    Set includer = CreateObject( "VBScripting.Includer" )
    Execute includer.Read( "VBSLogger" )
    Set logger = New VBSLogger
    Set sh = CreateObject( "WScript.Shell" )
    response = sh.PopUp("message", 60, "caption", vbOKCancel) '1st dialog
    MsgBox response '2nd dialog
    Set sh = Nothing
    Set includer = Nothing
    ```
- Note that the `Set includer = Nothing` statement is called after all PopUp method calls.

<br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
<br /><br /><br /><br /><br /><br /><br /><br /><br /><br />