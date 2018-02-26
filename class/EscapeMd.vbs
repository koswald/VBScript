'EscapeMd.vbs

'Escapes markdown special characters.

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