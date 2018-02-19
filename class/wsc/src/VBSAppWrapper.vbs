Dim app, WScript, document
Set app = New VBSApp

'Property Init
'Parameter: an object
'Returns: a self reference
'Remark: Pass in the WScript object from a script or the document object from an .hta: Set app = CreateObject("VBScripting.VBSApp") : Init(WScript)
Function Init(obj)
    Dim errMsg : errMsg = "Expected .vbs/.wsf WScript object or .hta document object." & vbLf & "Actual: " & TypeName(obj)
    Dim errSrc : errSrc = "VBSApp.wsc, Init Function"
    If "HTMLDocument" = TypeName(obj) Then
        Set document = obj
    ElseIf "Object" = TypeName(obj) Then
        Dim erred : erred = False
        On Error Resume Next
            Dim x : x = obj.ScriptName
            If Err Then erred = True
        On Error Goto 0
        If erred Then Err.Raise 2, errSrc, errMsg
        Set WScript = obj
    Else
        Err.Raise 1, errSrc, errMsg
    End If
    app.InitializeAppTypes
    Set Init = Me
End Function

Sub Close
    Set WScript = Nothing
    Set document = Nothing
    Set app = Nothing
End Sub

'Wrap the VBSApp class public members
        
Function GetArgs : GetArgs = app.GetArgs : End Function
Function GetArgsString : GetArgsString = app.GetArgsString : End Function
Function GetArg(index) : GetArg = app.GetArg(index) : End Function
Function GetArgsCount : GetArgsCount = app.GetArgsCount : End Function
Function GetFullName : GetFullName = app.GetFullName : End Function
Function GetFileName : GetFileName = app.GetFileName : End Function
Function GetBaseName : GetBaseName = app.GetBaseName : End Function
Function GetExtensionName : GetExtensionName = app.GetExtensionName : End Function
Function GetParentFolderName : GetParentFolderName = app.GetParentFolderName : End Function
Function GetExe : GetExe = app.GetExe : End Function
Sub RestartWith(host, switch, elevating) : app.RestartWith host, switch, elevating : End Sub
Sub SetUserInteractive(newUserInteractive) : app.SetUserInteractive newUserInteractive : End Sub
Function GetUserInteractive : GetUserInteractive = app.GetUserInteractive : End Function
Sub SetVisibility(newVisibility) : app.SetVisibility newVisibility : End Sub
Function GetVisibility : GetVisibility = app.GetVisibility : End Function
Sub Quit : app.Quit : End Sub
Sub Sleep(milliseconds) : app.Sleep milliseconds : End Sub
Function WScriptHost : WScriptHost = app.WScriptHost : End Function
Function CScriptHost : CScriptHost = app.CScriptHost : End Function
Function GetHost : GetHost = app.GetHost : End Function
Function GetWrapAll : GetWrapAll = app.WrapAll : End Function
Sub PutWrapAll(newWrapAll) : app.WrapAll = newWrapAll : End Sub
