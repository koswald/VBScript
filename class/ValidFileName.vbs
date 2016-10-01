
'Provides for removing characters in a string that are not suitable for use in a Windows&reg file name.

'Example of well-formed comments in a script without a Class statement:
'A line starting with '''' (four single quotes) separates the general comments at the beginning of the script from the rest of the script

'Usage Example
''
''    With CreateObject("includer")
''        ExecuteGlobal(.read("ValidFileName"))
''    End With
''
''    MsgBox GetValidFileName("test\ing") 'test-ing
'

'''' End general comments. For the DocGenerator, this line takes the place of a Class statement.

'Function GetValidFileName
'Parameter: a file name candidate
'Returns a valid file name
'Remarks: Returns a string suitable for use as a file name: Strips ("\", "/", ":", "*", "?", """", "<", ">", "|", "%20") and replaces them with a hyphen/dash (-)

Function GetValidFileName(fileNameCandidate)

   'items #1 thru #9: a Windows file name can't contain any of these
   'item #10: Chrome won't open an .html file with %20 in the title
   invalidItems = Array("\", "/", ":", "*", "?", """", "<", ">", "|", "%20")

   'a file name has a max value, somewhere in the neightborhood of 200 characters
   Const maxLength = 130

   x = fileNameCandidate

   x = Left(x, maxLength)

   For i = 0 to UBound(invalidItems)

      x = Replace(x, invalidItems(i), "-")

   Next
   GetValidFileName = x
End Function