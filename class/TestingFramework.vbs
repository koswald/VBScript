'A lightweight testing framework

'Usage example
' <pre>     With CreateObject( "VBScripting.Includer" ) <br />         Execute .Read( "VBSValidator" ) <br />         Execute .Read( "TestingFramework" ) <br />     End With <br />     With New TestingFramework <br />         .Describe "VBSValidator class" <br />             Dim val : Set val = New VBSValidator 'class under test <br />         .It "should return False when IsBoolean is given a string" <br />             .AssertEqual val.IsBoolean( "sdfjke" ), False <br />         .It "should raise an error when EnsureBoolean is given a string" <br />             Dim nonBool : nonBool = "a string" <br />             On Error Resume Next <br />                 val.EnsureBoolean(nonBool) <br />                 .AssertErrorRaised <br />                 Dim errDescr : errDescr = Err.Description<br />                 Dim errSrc : errSrc = Err.Source <br />             On Error Goto 0 <br />     End With </pre>
'
' When a test file such as <code>spec\Configurer.spec.wsf</code> is double-clicked in Windows Explorer, the default Windows behavior is to open the script with wscript.exe, but the test requires cscript.exe, so the file is automatically restarted with cscript.exe. By default, the test opens with PowerShell in Windows Terminal, if installed. This behavior may be changed by adding a "shell" key/value pair to <code>class\VBSHoster.configure</code>, overriding the default behavior. Alternatively, a script-specific .configure file can be added; see the <a href="#configurer"> Configurer class docs</a>. 
'
' See also <a href="#vbstestrunner"> VBSTestRunner</a> and <a href="#vbshoster"> VBSHoster</a>.
'
Class TestingFramework

    Private unit 'string: description of the unit under test
    Private spec 'string: description of a specification or test or expectation
    Private T 'a string of spaces (tab)
    Private explanation 'string: reason for test failure
    Private pass 'string literal: "Pass"
    Private fail 'string literal: "Fail"
    Private result '"Pass" or "Fail"
    Private resultPending 'boolean
    Private sh 'WScript.Shell object
    Private fso 'Scripting.FileSystemObject object
    Private sendKeysWarning 'process returned by the Exec method: a WScript MsgBox

    Private Sub Class_Initialize
        Dim hoster 'VBSHoster object
        Set sh = CreateObject( "WScript.Shell" )
        Set fso = CreateObject( "Scripting.FileSystemObject" )
        sh.CurrentDirectory = fso.GetParentFolderName( WScript.ScriptFullName)
        SetResultPending False
        pass = "Pass"
        fail = "Fail"
        T = "      "
        With CreateObject( "VBScripting.Includer" )
            Execute .Read( "VBSHoster" )
            Set hoster = New VBSHoster
            hoster.EnsureCScriptHost 'allow file double-click in explorer to run a test
        End With
    End Sub

    Private Sub WriteLine(str)
        If Len(str) Then WScript.StdOut.WriteLine str
    End Sub

    'Method Describe
    'Parameter: unit description
    'Remark: Sets the description for the unit under test. E.g. .describe "DocGenerator class"
    Sub Describe(newUnit)
        ShowPendingResult
        unit = newUnit
        If Len(newUnit) Then WriteLine Left("--------- " & newUnit & " ---------------------------------------------------------", 79)
    End Sub

    'Method It
    'Parameter: an expectation
    'Remark: Sets the specification, a.k.a. spec, which is a description of some expectation to be met by the unit under test. E.g. .it "should return an integer"
    Sub It(newSpec)
        ShowPendingResult
        spec = newSpec
    End Sub

    'Property GetSpec
    'Returns a string
    'Remark: Returns the specification string for the current spec.
    Property Get GetSpec : GetSpec = spec : End Property

    'Method ShowPendingResult
    'Remark: Flushes any pending results. Generally for internal use, but may occasionally be helpful prior to an ad hoc StdOut comment, so that the comment shows up in the output in its proper place.
    Sub ShowPendingResult
        If Not resultPending Then Exit Sub
        WriteLine result & T & spec
        If fail = result Then
            If Len(explanation) Then WriteLine "========> " & explanation
        End If
        SetResultPending False
    End Sub

    'Method AssertEqual
    'Parameters: actual, expected
    'Remark: Asserts that the specified two variants, of any subtype, are equal.
    Sub AssertEqual(var1, var2)
        ShowPendingResult
        If var1 = var2 Then
            SetResult pass
        Else SetResult fail
            explanation = _
                "Expected: " & var2 & vbCrLf & _
                "========> " & _
                "Actual  : " & var1
        End If
        SetResultPending True
    End Sub

    'Method AssertErrorRaised
    'Remark: Asserts that an error should be raised by one or more of the preceeding statements. The statement(s), together with the AssertErrorRaised statement, should be wrapped with an <br /> <pre style='white-space: nowrap;'> On Error Resume Next <br /> On Error Goto 0 </pre> block.
    Sub AssertErrorRaised
        ShowPendingResult
        If Err Then
            SetResult pass
        Else
            SetResult fail
            explanation = "Expected error to be raised. Actual: no error"
        End If
        SetResultPending True
    End Sub

    'Method DeleteFile
    'Parameter: a filespec
    'Remark: Deletes the specified file. Relative paths and environment variables are allowed.
    Sub DeleteFile(file)
        fso.DeleteFile fso.GetAbsolutePathName(sh.ExpandEnvironmentStrings(file))
    End Sub

    'Method DeleteFiles
    'Parameter: an array
    'Remark: Deletes the specified files. The parameter is an array of filespecs. Relative paths and environment variables are allowed.
    Sub DeleteFiles(files)
        Dim file
        For Each file In files
            DeleteFile file
        Next
    End Sub

    'Method WriteTempMessage
    'Parameter: a string
    'Remark: Writes a temporary message to the test output that can be, and should be, erased later with the EraseTempMessage method, after some behind the scenes work has been done that does not write to the console. Note: The message will not appear when the test(s) are initiated by the TestRunner class.
    Sub WriteTempMessage( str )
        tempMessage_ = str
        ShowPendingResult
        WScript.StdOut.Write str
    End Sub
    Private tempMessage_

    'Method EraseTempMessage
    'Remarks: Erases the message written by the WriteTempMessage method.
    Sub EraseTempMessage
        Dim i
        For i = 1 To Len(tempMessage_)
            WScript.StdOut.Write Chr(8) & " " & Chr(8) 'backspace space backspace
            WScript.Sleep 1
        Next
    End Sub


    'Function MessageAppeared
    'Parameter: caption, seconds, keys
    'Returns: a boolean
    'Remark: Waits for the specified maximum time (seconds) for a dialog with the specified title-bar text (caption). If the dialog appears, acknowleges it with the specified keystrokes (keys) and returns True. If the time elapses without the dialog appearing, returns False. Note: SendKeys-related features are deprecated.
    Function MessageAppeared(caption, seconds, keys)
        Dim i : i = 0
        While (Not sh.AppActivate(caption)) And i < seconds * 250
            WScript.Sleep 4
            i = i + 1
        Wend
        If sh.AppActivate(caption) Then
            sh.SendKeys keys
            MessageAppeared = True
        Else MessageAppeared = False
        End If
    End Function

    Private Sub SetResult(newResult)
        result = newResult
    End Sub

    Private Sub SetResultPending(pBool)
        resultPending = pBool
        If Not resultPending Then
            explanation = ""
        End If
    End Sub

    'Method ShowSendKeysWarning
    'Remark: Shows a SendKeys warning: a warning message to not make mouse clicks or key presses. Note: SendKeys-related features are deprecated.
    Sub ShowSendKeysWarning
        With CreateObject( "VBScripting.Includer" )
            Execute .Read( "StringFormatter" )
            Set sendKeysWarning = sh.Exec((New StringFormatter)(Array( _
                "wscript ""%s\TestingFramework.fixture.vbs"" ""%s""", _
                .LibraryPath, WScript.ScriptName _
            )))
        End With
    End Sub

    'Method CloseSendKeysWarning
    'Remark: Closes the SendKeys warning. Note: SendKeys-related features are deprecated.
    Sub CloseSendKeysWarning
        sendKeysWarning.Terminate
    End Sub

    Sub Class_Terminate
        ShowPendingResult
        Set sh = Nothing
        Set fso = Nothing
    End Sub
End Class
