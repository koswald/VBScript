'Test the CommandParser class

Option Explicit
Dim cp 'CommandParser object: object in test
Dim incl 'VBScripting.Includer object
Dim OsBitnessCmd 'string
Dim actual, expected
Dim echoHelpCmd, echoHelpOut 'strings
Dim searchString, searchString2 'strings
Dim PowerShellCmd 'string

Set incl = CreateObject( "VBScripting.Includer" )
OsBitnessCmd = "cmd /c If defined ProgramFiles^(X86^) (echo 64-bit) else (echo 32-bit)"
echoHelpCmd = "cmd /c echo /?"
echoHelpOut = "Displays messages, or turns command-echoing on or off." & vbCrLf & vbCrLf & "  ECHO [ON | OFF]" & vbCrLf & "  ECHO [message]" & vbCrLf & vbCrLf & "Type ECHO without parameters to display the current echo setting." & vbCrLf
PowerShellCmd = "powershell -NonInteractive -ExecutionPolicy Bypass -NoProfile -Command (Get-Service Spooler).Status.Value__"

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "CommandParser class"
        Set cp = incl.LoadObject( "CommandParser" )

    .It "should get single-line output"
        actual = cp.Out( OsBitnessCmd )
        expected = "64-bit" & vbCrLf
        .AssertEqual actual,  expected

    .It "should get multi-line output"
        actual = cp.Out( echoHelpCmd )
        expected = echoHelpOut
        .AssertEqual actual, expected

    .It "should search multi-line output (positive)"
        searchString = "t parameters"
        actual = cp.Srch( echoHelpOut, searchString )
        expected = True
        .AssertEqual actual, expected

    .It "should combine the above functions"
        actual = cp.Result( echoHelpCmd, searchString )
        expected = True
        .AssertEqual actual, expected

    .It "should search multi-line output (negative)"
        searchString2 = "t paramaters" 'misspelling
        actual = cp.Srch( echoHelpOut, searchString2 )
        expected = False
        .AssertEqual actual, expected

    .It "should replace vbCrLf with vbLf"
        actual = cp.ReplaceCrLf( echoHelpOut )
        expected = "Displays messages, or turns command-echoing on or off." & vbLf & vbLf & "  ECHO [ON | OFF]" & vbLf & "  ECHO [message]" & vbLf & vbLf & "Type ECHO without parameters to display the current echo setting." & vbLf
        .AssertEqual actual, expected

    .It "should get text from a PowerShell process"
        .WriteTempMessage "Starting Windows PowerShell..."
        actual = CInt( Left( cp.Out( PowerShellCmd ), 1 )) 'Left strips off trailing vbCrLf
        If 0 = actual Then
            .EraseTempMessage
            Err.Raise 51,, """" & .GetSpec & """ errored: 0 is an invalid spooler service status."
        End If
        .EraseTempMessage
        expected = 1 Or 4 '1 = stopped, 4 = running
        .AssertEqual actual Or 1 Or 4, expected
End With
