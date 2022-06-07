'Test the VBScripting.VBSPower Windows Script Component

Option Explicit
Dim pwr 'VBScripting.VBSPower object
Dim incl 'VBScripting.Includer object

Set incl = CreateObject( "VBScripting.Includer" )

Execute incl.Read( "TestingFramework" )
With New TestingFramework

    .describe "VBSPower class"
        Set pwr = CreateObject( "VBScripting.VBSPower" )

    'setup
        pwr.SetDebug True 'True => don't actually power down, restart, etc

    ' Calling the following methods without producing an error shows that a method with that name exists.

    .it "should power down the computer"
        .AssertEqual TypeName(pwr.Shutdown), "Empty"

    .it "should restart the computer"
        .AssertEqual TypeName(pwr.Restart), "Empty"

    .it "should log off of the user session"
        .AssertEqual TypeName(pwr.Logoff), "Empty"

    .it "should put the computer to sleep"
        .AssertEqual TypeName(pwr.Sleep), "Empty"

    .it "should put the computer into hibernation"
        .AssertEqual TypeName(pwr.Hibernate), "Empty"

    .it "should enable hibernation"
        .AssertEqual TypeName(pwr.EnableHibernation), "Empty"

    .it "should disable hibernation"
        .AssertEqual TypeName(pwr.DisableHibernation), "Empty"

End With

