
'Command Parser
'
'Runs a specified command and searches the output for a phrase

Class CommandParser

    Private cmd, startPhrase, stopPhrase, searchPhrase
    Private sh

    Sub Class_Initialize
        Set sh = CreateObject("WScript.Shell")
        With CreateObject("includer")
            Execute .read("GUIDGenerator")
        End With
        Dim newGuid : Set newGuid = New GUIDGenerator 'provides a unique string
        SetCommand ""
        SetStartPhrase ""
        SetStopPhrase newGuid
        SetSearchPhrase newGuid
    End Sub
    
    'Method SetCommand
    'Parameter: newCmd
    'Remark: Sets the command to run whose output will be analyzed/parsed.
    Sub SetCommand(newCmd) : cmd = newCmd : End Sub

    'Method SetSearchPhrase
    'Parameter: newSearchPhrase
    'Remark: Sets a phase to search for in the command's output
    Sub SetSearchPhrase(newSearchPhrase) : searchPhrase = newSearchPhrase : End Sub

    Property Get GetCommand : GetCommand = cmd : End Property
    Property Get GetStartPhrase : GetStartPhrase = startPhrase : End Property
    Property Get GetStopPhrase : GetStopPhrase = stopPhrase : End Property
    Property Get GetSearchPhrase : GetSearchPhrase = searchPhrase : End Property

    'Property GetResult
    'Returns: a boolean
    'Remark: Runs the sepecified command and returns True if the specified phrase is found in the command output.
    Property Get GetResult
        GetResult = False
        Dim pipe : Set pipe = sh.Exec(cmd)
        Dim line : line = ""
        Dim started : If startPhrase = "" Then started = True Else started = False
        Dim stopped : stopped = False
        While Not pipe.StdOut.AtEndOfStream
            line = pipe.StdOut.ReadLine
            If InStr(line, stopPhrase) Then stopped = True
            If InStr(line, startPhrase) Then started = True
            If started And Not stopped Then If InStr(line, searchPhrase) Then GetResult = True
        Wend
        On Error Resume Next
            pipe.Terminate
        On Error Goto 0
        Set pipe = Nothing
    End Property

    'Method SetStartPhrase
    'Parameter: newStartPhrase
    'Remark: Sets a unique phrase to identify the output line after which the search begins. Optional. By defualt the output is searched from the beginning.
    Sub SetStartPhrase(newStartPhrase) : startPhrase = newStartPhrase : End Sub

    'Method SetStopPhrase
    'Parameter: newStopPhrase
    'Remark: Sets a unique phrase to identify the line that follows the last line of the search. Optional. By defualt, the output is searched to the end.
    Sub SetStopPhrase(newStopPhrase) : stopPhrase = newStopPhrase : End Sub

    Sub Class_Terminate
        Set sh = Nothing
    End Sub

End Class
