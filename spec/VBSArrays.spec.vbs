'Test the VBSArrays class

Option Explicit
Dim va
Dim incl

Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .describe "VBSArrays class"
        Execute incl.Read( "VBSArrays" )
        Set va = New VBSArrays

    .it "should return an array without duplicate values"
        .AssertEqual Join( va.Uniques(Split("str str" ))), "str"

End With
