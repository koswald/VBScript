'EscapeMd and EscapeMd2 Functions

'Escapes markdown special characters.
'
''''

'Function EscapeMd
'Parameters: unescaped string
'Returns: escaped string
'Remarks: Returns a string with Markdown special characters escaped.
Function EscapeMd(ByVal str) 'conversions for markdown
    str = Replace(str, "\", "\\")
    str = Replace(str, "`", "\`")
    str = Replace(str, "*", "\*")
    str = Replace(str, "_", "\_")
    str = Replace(str, "{", "\{")
    str = Replace(str, "}", "\}")
    str = Replace(str, "[", "\[")
    str = Replace(str, "]", "\]")
    str = Replace(str, "(", "\(")
    str = Replace(str, ")", "\)")
    str = Replace(str, "#", "\#")
    str = Replace(str, "+", "\+")
    str = Replace(str, "-", "\-")
    str = Replace(str, ".", "\.")
    str = Replace(str, "!", "\!")
    EscapeMd = Replace(str, "|", "\|")
End Function

'From https://meta.stackexchange.com/questions/82718/how-do-i-escape-a-backtick-in-markdown#198231
' \   backslash
' `   backtick
' *   asterisk
' _   underscore
' {}  curly braces
' []  square brackets
' ()  parentheses
' #   hash mark
' +   plus sign
' -   minus sign (hyphen)
' .   dot
' !   exclamation mark

' to which I would add | for tables;
' also, make sure to replace \ with \\ first of all

'Function EscapeMd2
'Parameters: unescaped string
'Returns: escaped string
'Remarks: Returns a string with a minimal amount of Markdown special characters escaped. <a href="http://www.theukwebdesigncompany.com/articles/entity-escape-characters.php"> Escape codes</a>.
Function EscapeMd2(str)
    Dim s : s = str
    s = Replace(s, "|", "&#124;")
    s = Replace(s, "*", "&#42;")
    EscapeMd2 = s
End Function