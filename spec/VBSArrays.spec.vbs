
'Test the VBSArrays class

With CreateObject("includer")
    Execute(.read("VBSArrays"))
    Execute(.read("TestingFramework"))
End With

Dim va : Set va = New VBSArrays

With New TestingFramework

    .describe "VBSArrays class"

        .it "should return an array without duplicate values"
        
            .AssertEqual Join(va.Uniques(Split("str str"))), "str"
End With