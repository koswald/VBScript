Dim WScript
Dim document
Dim app_X
Set app_X = New VBSApp

Sub Close
    Set app_X = Nothing
End Sub

'Method: Init
'Parameter: an object
'Returns: a self reference
'Remark: Pass in the WScript object from a .vbs or .wsf script or pass in the Document object from an .hta application when the class is instantiated as follows: Set app = CreateObject( "VBScripting.VBSApp" ) : app.Init WScript. The Init method is *not* required when the VBSApp class is included by 1) direct reference in an .hta or .wsf file <script> tag in the src attribute; or 2) by an Execute includer.Read( "ClassName" ) statement.
Sub Init(obj)
    If "HTMLDocument" = TypeName(obj) Then
        Set document  = obj
    ElseIf "Object" = TypeName(obj) Then
        Set WScript = obj
    End If
    app_X.InitializeAppTypes
End Sub

'Wrap the VBSApp class public members

Function GetArgs
    GetArgs = app_X.GetArgs
End Function
Function GetArgsString
    GetArgsString = app_X.GetArgsString
End Function
Function GetArg(index)
    GetArg = app_X.GetArg(index)
End Function
Function GetArgsCount
    GetArgsCount = app_X.GetArgsCount
End Function
Function GetFullName
    GetFullName = app_X.GetFullName
End Function
Function GetFileName
    GetFileName = app_X.GetFileName
End Function
Function GetBaseName
    GetBaseName = app_X.GetBaseName
End Function
Function GetExtensionName
    GetExtensionName = app_X.GetExtensionName
End Function
Function GetParentFolderName
    GetParentFolderName = app_X.GetParentFolderName
End Function
Function GetExe
    GetExe = app_X.GetExe
End Function
Sub RestartWith(host, switch, elevating)
    app_X.RestartWith host, switch, elevating
End Sub
Sub SetUserInteractive(newUserInteractive)
    app_X.SetUserInteractive newUserInteractive
End Sub
Function GetUserInteractive
    GetUserInteractive = app_X.GetUserInteractive
End Function
Sub SetVisibility(newVisibility)
    app_X.SetVisibility newVisibility
End Sub
Function GetVisibility
    GetVisibility = app_X.GetVisibility
End Function
Sub Quit
    app_X.Quit
End Sub
Sub Sleep(milliseconds)
    app_X.Sleep milliseconds
End Sub
Function WScriptHost
    WScriptHost = app_X.WScriptHost
End Function
Function CScriptHost
    CScriptHost = app_X.CScriptHost
End Function
Function GetHost
    GetHost = app_X.GetHost
End Function
Function GetWrapAll
    GetWrapAll = app_X.WrapAll
End Function
Sub PutWrapAll(newWrapAll)
    app_X.WrapAll = newWrapAll
End Sub


