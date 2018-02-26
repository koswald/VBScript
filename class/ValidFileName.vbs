
'Provides for modifying a string to remove characters that are not suitable for use in a Windows&reg; file name.

'Usage Example
'<pre>     With CreateObject("VBScripting.Includer") <br />         Execute .read("ValidFileName") <br />     End With <br />  <br />     MsgBox GetValidFileName("test\ing") 'test-ing </pre>
'

'ValidFileName.vbs provides an example of introductory comments in a script that lacks a Class statement: With DocGenerator.vbs, a line beginning with '''' (four single quotes) may be used instead of a Class statement, in order to end the introductory comments section.

'''' End general comments

'Function GetValidFileName
'Parameter: a file name candidate
'Returns a valid file name
'Remarks: Returns a string suitable for use as a file name: Removes <strong> \ / : * ? " < > | %20 # </strong> and replaces them with a hyphen/dash (-)

Function GetValidFileName(fileNameCandidate)

    'items 1 - 9: a Windows file name can't contain any of these
    'items 10 - 11: Chrome won't open a local .html file with %20 or # in the title
    Dim invalidItems : invalidItems = Array("\", "/", ":", "*", "?", """", "<", ">", "|", "%20", "#")

    Dim maxLength
    With CreateObject("VBScripting.Includer")
        Execute(.read("ValidFileName.config")) 'get the maxLength variable
    End With

    Dim x : x = fileNameCandidate

    Dim i
    For i = 0 to UBound(invalidItems)

        x = Replace(x, invalidItems(i), "-")

    Next
    GetValidFileName = Left(x, maxLength)
End Function

