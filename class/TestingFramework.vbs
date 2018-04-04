
'A lightweight testing framework

'Usage example
' <pre>     With CreateObject("VBScripting.Includer") <br />         Execute .read("VBSValidator") <br />         Execute .read("TestingFramework") <br />     End With <br />     Dim val : Set val = New VBSValidator 'class under test <br />     With New TestingFramework <br />         .describe "VBSValidator class" <br />         .it "should return False when IsBoolean is given a string" <br />             .AssertEqual val.IsBoolean("sdfjke"), False <br />         .it "should raise an error when EnsureBoolean is given a string" <br />             Dim nonBool : nonBool = "a string" <br />             On Error Resume Next <br />                 val.EnsureBoolean(nonBool) <br />                 .AssertErrorRaised <br />                 Dim errDescr : errDescr = Err.Description 'capture the error information <br />                 Dim errSrc : errSrc = Err.Source <br />             On Error Goto 0 <br />     End With </pre>
'
' See also VBSTestRunner
'
Class TestingFramework

    Private Sub WriteLine(str)
        If Len(str) Then WScript.StdOut.WriteLine str
    End Sub

    'Method describe
    'Parameter: unit description
    'Remark: Sets the description for the unit under test. E.g. .describe "DocGenerator class"
    Sub describe(newUnit)
        ShowPendingResult
        unit = newUnit
        If Len(unit) Then WriteLine Left("--------- " & newUnit & " ---------------------------------------------------------", 79)
    End Sub

    'Method it
    'Parameter: an expectation
    'Remark: Sets the specification, a.k.a. spec, which is a description of some expectation to be met by the unit under test. E.g. .it "should return an integer"
    Sub it(newSpec)
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
        Else
            SetResult fail
            explanation = "Expected: " & var2 & "; Actual: " & var1
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

    'Method DeleteFiles
    'Parameter: an array
    'Remark: Deletes the specified files. The parameter is an array of filespecs. Relative paths may be used.
    Sub DeleteFiles(files)
        Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
        Dim file
        For each file in files
            If fso.FileExists(file) Then fso.DeleteFile(file)
        Next
        Set fso = Nothing
    End Sub

    'Function MessageAppeared
    'Parameter: caption, seconds, keys
    'Returns: a boolean
    'Remark: Waits for the specified maximum time (seconds) for a dialog with the specified title-bar text (caption). If the dialog appears, acknowleges it with the specified keystrokes (keys) and returns True. If the time elapses without the dialog appearing, returns False.
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
    'Remark: Shows a SendKeys warning: a warning message to not make mouse clicks or key presses.
    Sub ShowSendKeysWarning
        With CreateObject("VBScripting.Includer")
            Execute .read("StringFormatter")
	        Set sendKeysWarning = sh.Exec((New StringFormatter)(Array( _
                "wscript ""%s\TestingFramework.fixture.vbs"" ""%s""", _
                .LibraryPath, WScript.ScriptName _
            )))
        End With
    End Sub

    'Method CloseSendKeysWarning
    'Remark: Closes the SendKeys warning.
    Sub CloseSendKeysWarning
        sendKeysWarning.Terminate
    End Sub

    Private unit, spec, T, explanation
    Private pass, fail, result, resultPending
    Private sh
    Private sendKeysWarning

    Private Sub Class_Initialize 'event fires on object instantiation
        SetResultPending False
        pass = "Pass" : fail = "Fail" : T = "      "
        With CreateObject("VBScripting.Includer")
            Execute .Read("VBSHoster")
            Dim hoster : Set hoster = New VBSHoster
            hoster.EnsureCScriptHost 'allow file double-click in explorer to run a test
        End With
        Set sh = CreateObject("WScript.Shell")
    End Sub

    Sub Class_Terminate
        ShowPendingResult
        Set sh = Nothing
    End Sub
End Class
