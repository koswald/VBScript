'Integration test for the WMIUtility class

Option Explicit
Dim wu 'WMIUtility object
Dim incl 'VBScripting.Includer object
Dim actual, expected
Dim tname 'TypeName return value
Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .Describe "WMIUtility class"
        Execute incl.Read( "WMIUtility" )
        Set wu = New WMIUtility

    .It "should get a battery object"
        actual = TypeName(wu.Battery)
        expected = "SWbemObjectEx"
        .AssertEqual actual, expected

    .It "should get the estimated charge remaining"
        tname = TypeName(wu.Battery.EstimatedChargeRemaining)
        actual = "Integer" = tname Or "Null" = tname
        expected = True
        .AssertEqual actual, expected

End With
