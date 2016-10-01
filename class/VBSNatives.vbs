
'Wrap objects native to WScript/Windows

Class VBSNatives

    Private oSh, oFSO, oArgs, oSA

    Sub Class_Initialize
        Set oSh = CreateObject("WScript.Shell")
        Set oFSO = CreateObject("Scripting.FileSystemObject")
        On Error Resume Next 'for including this file in .wsc files, where WScript is not available by default
            Set oArgs = WScript.Arguments
        On Error Goto 0
        Set oSA = CreateObject("Shell.Application")
    End Sub

    'Property shell
    'Returns a reference to a WScript.Shell object instance
    Property Get shell : Set shell = oSh : End Property
    'Property sh
    'Returns a reference to a WScript.Shell object instance
    Property Get sh : Set sh = oSh : End Property 'shell shortcut method

    'Property fso
    'Returns a reference to a Scripting.FileSystemObject object instance
    Property Get fso : Set fso = oFSO : End Property

    'Property args
    'Returns the WScript.Arguments collection
    Property Get args : Set args = oArgs : End Property
    'Property a
    'Returns the WScript.Arguments collection
    Property Get a : Set a = oArgs : End Property

    'Property dict
    'Returns a new Scripting.Dictionary object
    Property Get dict : Set dict = CreateObject("Scripting.Dictionary") : End Property 'get a shiny new dictionary object

    'Property ShellApp
    'Returns a reference to a Shell.Application object instance
    Property Get ShellApp : Set ShellApp = oSA : End Property
    'Property sa
    'Returns a reference to a Shell.Application object instance
    Property Get sa : Set sa = oSA : End Property

    Sub Class_Terminate 'event fires when the object instance goes out of scope
        Set oSh = Nothing
        Set oFSO = Nothing
        Set oArgs = Nothing
        Set oSA = Nothing
    End Sub

End Class
