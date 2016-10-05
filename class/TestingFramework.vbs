
'A lightweight testing framework

'Usage example
'
''    With CreateObject("includer")
''        Execute(.read("VBSValidator")) 'ExecuteGlobal if this code is wrapped in a function, class, etc.
''        Execute(.read("TestingFramework"))
''    End With
'' 
''    Dim val : Set val = New VBSValidator 'Class Under Test
'' 
''    With New TestingFramework
'' 
''        .describe "VBSValidator class"
'' 
''        .it "should return False when IsBoolean is given a string"
'' 
''            .AssertEqual val.IsBoolean("sdfjke"), False
'' 
''        .it "should raise an error when EnsureBoolean is given a string"
'' 
''            Dim nonBool : nonBool = "a string"
''            On Error Resume Next
''                val.EnsureBoolean(nonBool)
'' 
''                .AssertErrorRaised
'' 
''                Dim errDescr : errDescr = Err.Description 'capture the error information
''                Dim errSrc : errSrc = Err.Source
''            On Error Goto 0
''    End With
'

Class TestingFramework

    Private unit, spec, T, explanation
    Private pass, fail, result, resultPending

    Private Sub Class_Initialize 'event fires on object instantiation
        SetResultPending False
        pass = "Pass" : fail = "Fail" : T = "      "
    End Sub

    Private Sub WriteLine(str)
        If Len(str) Then WScript.StdOut.WriteLine str
    End Sub

    'Method describe
    'Parameter: unit description
    'Remark: Provides a description for the unit under test. E.g. describe "DocGenerator class"

    Sub describe(newUnit)
        ShowPendingResult
        unit = newUnit
        If Len(unit) Then WriteLine Left("--------- " & newUnit & " ---------------------------------------------------------", 80)
    End Sub

    'Method it
    'Parameter: an expectation
    'Remark: Provides a description of some expectation to be met by the unit under test. E.g. it "should return an integer"

    Sub it(newSpec)
        ShowPendingResult
        spec = newSpec
    End Sub

    Private Sub ShowPendingResult
        If Not resultPending Then Exit Sub
        WriteLine result & T & spec
        If fail = result Then
            If Len(explanation) Then WriteLine "========> " & explanation
        End If
        SetResultPending False
    End Sub

    'Method AssertEqual
    'Parameters: variant1, variant2
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
    'Remark: Asserts that an error should be raised by one or more of the preceeding statements. The statement(s), together with the AssertErrorRaised statement, should be wrapped with a <br /> <pre> On Error Resume Next <br /> On Error Goto 0 </pre> block.

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

    Private Sub SetResult(newResult)
        result = newResult
    End Sub

    Private Sub SetResultPending(pBool)
        resultPending = pBool
        If Not resultPending Then
            explanation = ""
        End If
    End Sub

    Sub Class_Terminate
        ShowPendingResult
    End Sub

End Class
