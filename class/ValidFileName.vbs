'VBS function GetValidFileName and associated functions provide for modifying a string to remove characters that are not suitable for use in a Windows&reg; file name.

'Usage Example
'<pre>     With CreateObject("VBScripting.Includer") <br />         ExecuteGlobal .Read("ValidFileName") <br />     End With <br />  <br />     MsgBox GetValidFileName("test\ing") 'test-ing </pre>
'
'ValidFileName.vbs provides an example of introductory comments in a script that lacks a Class statement: With DocGenerator.vbs, a line beginning with '''' (four single quotes) may be used instead of a Class statement, in order to end the introductory comments section.
'
'''' End general comments

'Function GetValidFileName
'Parameter: a file name candidate
'Returns a valid file name
'Remarks: Returns a string suitable for use as a file name: Removes <strong> \ / : * ? " < > &#124; %20 # </strong> and replaces them with a hyphen/dash (-). Limits length to maxLength value in ValidFileName.config.
Function GetValidFileName(fileNameCandidate)
    Dim maxLength
    Dim includer : Set includer = CreateObject("VBScripting.Includer")
    Execute includer.Read("ValidFileName.config") 'get maxLength
    Set includer = Nothing
    Dim arr, i, x : x = fileNameCandidate
    arr = InvalidWindowsFilenameChars
    If "Empty" = TypeName(arr) Then Err.Raise 1,, "Use ExecuteGlobal, not Execute, with Function-based scripts like ValidFileName.vbs, when scope is not global."
    For i = 0 to UBound(arr)
        x = Replace(x, arr(i), "-")
    Next
    arr = InvalidChromeFilenameStrings
    For i = 0 to UBound(arr)
        x = Replace(x, arr(i), "-")
    Next
    GetValidFileName = Left(x, maxLength)
End Function

'Function InvalidWindowsFilenameChars
'Returns an array
'Remark: Returns an array of characters that are not allowed in Windows&reg; filenames.
Function InvalidWindowsFilenameChars : InvalidWindowsFilenameChars = Split("\ / : * ? "" < > |") : End Function

'Function InvalidChromeFilenameStrings
'Returns an array
'Remark: Returns an array of strings, either one of which if included in the filename of a local .html file, Chrome will not open the file.
Function InvalidChromeFilenameStrings : InvalidChromeFilenameStrings = Split("%20 #") : End Function
