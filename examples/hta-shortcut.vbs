'drop target
'extracts icon and descriptive info from the .hta
'and creates a shortcut to the .hta with that info

With WScript.Arguments
    If .Count = 0 Then Err.Raise 1,, "Argument required: an .hta file"
    hta = .item(0)
End With
Set fso = CreateObject("Scripting.FileSystemObject")
hta = fso.GetAbsolutePathName(hta) 'support relative path
If Not fso.FileExists(hta) Then Err.Raise 1,, "Couldn't find the file { " & hta & " }"
parent = fso.GetParentFolderName(hta)
base = fso.GetBaseName(hta)
With CreateObject("VBScripting.Includer")
    Execute .Read("VBSExtracter")
End With
With New VBSExtracter
    .SetIgnoreCase True

     'extract icon="word.word" up to and including the space following
     'there may or may not be whitespace around the =
     'there may or may not be quotes around the icon
    .SetPattern "icon\s*=\s*""?\w+\.{1}\w+""?\s*?"
    .SetFile hta
    icon = .extract

    'applicationName is similar, but in addition to word characters, 
    'there may be a . or -
    'others will have to be added later as they are encountered
    .SetPattern "applicationName\s*=\s*""?[\w-\.\s]+""?\s*?"
    name = .extract
End With
With New RegExp
    .IgnoreCase = True
    'setup to capture the match: word.word
    .Pattern = "icon\s*=\s*""?(\w+\.{1}\w+)""?"
    icon = .Replace(icon, "$1")
    .Pattern = "applicationName\s*=\s*""?([\w-\.\s]+)""?"
    name = .Replace(name, "$1")
End With
Set sh = CreateObject("WScript.Shell")
Set link = sh.CreateShortcut(parent & "\" & base & ".hta.lnk")
link.IconLocation = FindIcon(icon)
link.Arguments = ""
link.Description = name
link.WorkingDirectory = """" & parent & """"
link.TargetPath = """" & hta & """"
link.Save

Set fso = Nothing
Set sh = Nothing

Function Expand(str)
    Expand = sh.ExpandEnvironmentStrings(str)
End Function
Function Resolve(path)
    Resolve = fso.GetAbsolutePathName(path)
End Function
Function FindIcon(icon)
    folders = Array(".", "%SystemRoot%\System32", "%SystemRoot%")
    For Each folder in folders
        candidate = Resolve(Expand(folder)) & "\" & icon
        If fso.FileExists(candidate) Then
            FindIcon = candidate
            Exit Function
        End If
    Next
    Err.Raise 1,, "Couldn't find icon { " & icon & " }"
End Function
