'Provides for modifying a string to remove characters that are not suitable for use in a Windows&reg; file name.

'Usage Example
'<pre>     With CreateObject("VBScripting.Includer") <br />         Execute .Read("ValidFileName") <br />     End With <br />  <br />     MsgBox GetValidFileName("test\ing") 'test-ing </pre>
'
'ValidFileName.vbs provides an example of introductory comments in a script that lacks a Class statement: With DocGenerator.vbs, a line beginning with '''' (four single quotes) may be used instead of a Class statement, in order to end the introductory comments section.
'
'''' End general comments

Const invalidInFileName = "\_/_:_*_?_""_<_>_|"
Const invalidInChromeFileName = "%20_#" 'Chrome won't open a local .html file with %20 or # in the name

'Function GetValidFileName
'Parameter: a file name candidate
'Returns a valid file name
'Remarks: Returns a string suitable for use as a file name: Removes <strong> \ / : * ? " < > | %20 # </strong> and replaces them with a hyphen/dash (-). Limits length to maxLength value in ValidFileName.config.
Function GetValidFileName(fileNameCandidate)
    Dim maxLength
    Dim includer : Set includer = CreateObject("VBScripting.Includer")
    Execute includer.Read("ValidFileName.config") 'get maxLength
    Set includer = Nothing
    Dim arr, i, x : x = fileNameCandidate
    arr = Split(invalidInFileName, "_")
    For i = 0 to UBound(arr)
        x = Replace(x, arr(i), "-")
    Next
    arr = Split(invalidInChromeFileName, "_")
    For i = 0 to UBound(arr)
        x = Replace(x, arr(i), "-")
    Next
    GetValidFileName = Left(x, maxLength)
End Function
