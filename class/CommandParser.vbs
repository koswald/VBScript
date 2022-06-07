'The CommandParser class' Result method runs a command and searches its output for a phrase.
'
'Example:
'<pre> Dim includer : Set includer = CreateObject( "VBScripting.Includer" ) <br /> Execute includer.Read( "CommandParser" ) <br /> Dim cp : Set cp = New CommandParser <br /> Dim cmd : cmd = "cmd /c If defined ProgramFiles^(X86^) (echo 64-bit) else (echo 32-bit)" <br /> Dim phrase : phrase = "64-bit" <br /> MsgBox cp.Result( cmd, phrase ) 'True expected for 64-bit systems</pre>
'
Class CommandParser

    'Property Result
    'Returns: a boolean
    'Parameters: cmd, phrase
    'Remark: Runs the specified command and returns a boolean: True if the specified phrase is found in the output of the specified command. Not case sensitive by default.
    Property Get Result( byVal cmd, byVal phrase )
        Result = Srch( Out( cmd ), phrase )
    End Property

    'Testability Function Out
    'Returns: a string
    'Parameter: a command
    'Remarks: Returns the output of the specified Windows console command. May be multiple lines.
    Function Out( byVal cmd )
        With CreateObject( "WScript.Shell" )
           Out = .Exec( cmd ).StdOut.ReadAll
        End With
    End Function

    'Used solely for testing
    Function ReplaceCrLf( byVal str )
        ReplaceCrLf = Replace( str, vbCrLf, vbLf )
    End Function

    'Returns True if and only if str contains phrase.
    Function Srch( byVal str, byVal phrase )
        If CaseSensitive Then
            Srch = CBool( InStr( str, phrase ))
        Else Srch = CBool( InStr( LCase( str ), LCase( phrase )))
        End If
    End Function

    'Property CaseSensitive
    'Returns a boolean
    'Parameter: a boolean
    'Remark: Gets or sets whether the search is case sensitive. Default is False. 
    Property Get CaseSensitive
        CaseSensitive = caseSensitive_
     End Property
     Property Let CaseSensitive( newValue )
         caseSensitive_ = newValue
     End Property
     Private caseSensitive_
 
     Sub Class_Initialize
         CaseSensitive = False
     End Sub

End Class
