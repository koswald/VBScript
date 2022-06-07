
'Test ValidFileName.vbs

Option Explicit
Dim includer

Set includer = CreateObject( "VBScripting.Includer" )

Main

Sub Main
    Execute includer.Read( "TestingFramework" )
    With New TestingFramework

        .describe "ValidFileName.vbs"

        .it "should raise an error when Execute is used with non-global scope."
            Execute includer.Read( "ValidFileName" )
            On Error Resume Next
                Dim x : x = GetValidFileName( "xx" )
                .AssertEqual Err.Description, "Use ExecuteGlobal, not Execute, with Function-based scripts like ValidFileName.vbs, when scope is not global."
            On Error Goto 0

        .it "should return a string suitable for a filename"
            ExecuteGlobal includer.Read( "ValidFileName" )
            .AssertEqual GetValidFileName("\/:*?""<>|%20#"), "-----------"

        .it "should return characters invalid in a Windows filename"
            .AssertEqual Join(InvalidWindowsFilenameChars), "\ / : * ? "" < > |"

        .it "should return strings invalid to Chrome for a filename"
            .AssertEqual Join(InvalidChromeFilenameStrings), "%20 #"

    End With
End Sub
