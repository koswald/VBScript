
'Provides for modifying a string to remove characters that are not suitable for use in a Windows&reg; file name.

'Usage Example
''
''    With CreateObject("includer")
''        ExecuteGlobal(.read("ValidFileName"))
''    End With
''
''    MsgBox GetValidFileName("test\ing") 'test-ing
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
   invalidItems = Array("\", "/", ":", "*", "?", """", "<", ">", "|", "%20", "#")

   'a file name has a max value, somewhere in the neighborhood of 200 characters
   Const maxLength = 130

   x = fileNameCandidate

   x = Left(x, maxLength)

   For i = 0 to UBound(invalidItems)

      x = Replace(x, invalidItems(i), "-")

   Next
   GetValidFileName = x
End Function