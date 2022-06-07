
Option Explicit
Dim aoo 'ArrayOfObjects object: what is to be tested
Dim incl 'VBScripting.Includer
Dim actual, expected 'assertion arguments
Dim arr 'array
Dim x 'string

Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "ArrayOfObjects class"
        Execute incl.Read( "ArrayOfObjects" )
        Set aoo = New ArrayOfObjects

    .It "should return an array"
        actual = TypeName( aoo.Items )
        expected = "Variant()"
    .AssertEqual actual, expected

    .It "should return a zero-length array"
    .AssertEqual UBound( aoo.Items ), -1

    aoo.Add New TestTuple.Init( "n0", "val0" )
    
    .It "should get a property by index--syntax A: obj()(i).Value"
    .AssertEqual aoo()(0).Value, "val0"

    .It "should get a property by index--syntax B: obj.Items()(i).Value"
    .AssertEqual aoo.Items()(0).Value, "val0"
    
    .It "should get a property by index--syntax C: arr = obj.Items : arr(i).Value"
        arr = aoo.Items
    .AssertEqual arr(0).Value, "val0"
    
    .It "should get a property by index--syntax D: arr = obj : arr(i).Value"
        arr = aoo
    .AssertEqual arr(0).Value, "val0"

    .It "should not get a property by index--syntax X1"
        On Error Resume Next
            x = aoo.Items(0).Value
            actual = Err.Description
        On Error Goto 0
        expected = "Wrong number of arguments or invalid property assignment"
    .AssertEqual actual, expected

    .It "should not get a property by index--syntax X2"
        On Error Resume Next
            x = (aoo.Items)(0).Value
            actual = Err.Description
        On Error Goto 0
        expected = "Wrong number of arguments or invalid property assignment"
    .AssertEqual actual, expected

    .It "should not get a property by index--syntax X3"
        On Error Resume Next
            x = (aoo)(0).Value
            actual = Err.Description
        On Error Goto 0
        expected = "Wrong number of arguments or invalid property assignment"
    .AssertEqual actual, expected

    .It "should not get a property by index--syntax X4"
        On Error Resume Next
            x = aoo(0).Value
            actual = Err.Description
        On Error Goto 0
        expected = "Wrong number of arguments or invalid property assignment"
    .AssertEqual actual, expected
    
End With

Class TestTuple
    Public Key, Value
    Function Init( k, v )
        Key = k
        Value = v
        Set Init = me
    End Function
End Class