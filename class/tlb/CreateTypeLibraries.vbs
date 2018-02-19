'Create type libraries for selected .wsc files
Option Explicit : Initialize
For i = 1 To UBound(settings) Step 4
    AddURL format(Array( _
        "%s\wsc\%s.wsc", classPath, settings(i) _
    ))
    stl.GUID = settings(i + 1)
    stl.MajorVersion = settings(i + 2)
    stl.MinorVersion = settings(i + 3)
    stl.Name = settings(i) & "TLib"
    SetPath format(Array( _
        "%s\tlb\%s.tlb", classPath, settings(i) _
    ))
    On Error Resume Next
        stl.Write
    On Error Goto 0
    stl.Reset
Next
Set stl = Nothing
Set fso = Nothing

Sub AddURL(wscPath)
    If Not fso.FileExists(wscPath) Then Err.Raise 1,, "Couldn't find .wsc file" & vbLf & wscPath
    stl.AddURL wscPath
End Sub

Sub SetPath(tlbPath)
    If Not fso.FolderExists(fso.GetParentFolderName(tlbPath)) Then Err.Raise 2,, "Couldn't find parent folder of .tlb file" & vbLf & tlbPath
    stl.Path = tlbPath
End Sub

Dim settings, classPath, i
Dim fso, stl, format

Sub Initialize
    settings = Array("" _
        , "Includer" _
            , "{9C564563-8A4C-46FC-84F1-2E25F88CE784}" _
            , 1 _
            , 0 _
        , "StringFormatter" _
            , "{9C564563-8A4D-46FC-84F1-2E25F88CE784}" _
            , 1 _
            , 0 _
    )
    Set fso = CreateObject("Scripting.FileSystemObject")
    classPath = fso.GetAbsolutePathName("..")
    Set stl = CreateObject("Scriptlet.TypeLib")
    Set format = CreateObject("VBScripting.StringFormatter")
End Sub
