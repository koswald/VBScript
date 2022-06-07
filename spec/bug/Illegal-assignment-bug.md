# Illegal Assignment Bug

[Description](#description)  
[Steps to reproduce](#steps-to-reproduce)  
[Workarounds](#workarounds)

---

## Description

An *Illegal assignment* error may be raised when using the *New* keyword to assign an object instance for one of the project classes. 
Currently observed only when the variable name is `log`. 
This behavior is seen with other classes in the 'class' folder, not just VBSLogger, so it appears to be unrelated to the default method of the VBSLogger class being named Log.)

## Steps to reproduce

- Create a new .vbs file with the following code.
    ```vb
    With CreateObject( "VBScripting.Includer" )
        Execute .Read( "VBSLogger" )
        Set log = New VBSLogger
    End With
    ```
- Double click the .vbs file to run it.
- If an *Illegal assignment: log* error is received at line 3, then the bug is present.

## Workarounds

- Use a Dim statement:
    ```vb
    With CreateObject( "VBScripting.Includer" )
        Execute .Read( "VBSLogger" )
        Dim log : Set log = New VBSLogger
    End With
    ```
- Rename the variable.
    ```vb
    With CreateObject( "VBScripting.Includer" )
        Execute .Read( "VBSLogger" )
        Set logger = New VBSLogger
    End With
    ```
<br /><br /><br /><br /><br /><br /><br /><br /><br /><br />
<br /><br /><br /><br /><br /><br /><br /><br /><br /><br />