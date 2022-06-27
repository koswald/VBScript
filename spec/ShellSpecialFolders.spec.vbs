'Integration test for the ShellSpecialFolders class
'Use /ShowPaths to display the path for each folder

Option Explicit
Dim ssf 'ShellSpecialFolders object: what is to be tested
Dim incl 'VBScripting.Includer object
Dim sh 'WScript.Shell object
Dim actual, expected

Set incl = CreateObject("VBScripting.Includer")
Set sh = CreateObject("WScript.Shell")

Execute incl.Read("TestingFramework")
With New TestingFramework

    .Describe "ShellSpecialFolders class"
        Set ssf = incl.LoadObject("ShellSpecialFolders")

    .It "should get the path to the desktop"
        actual = ssf.Path( ssf.ssfDESKTOPDIRECTORY )
        expected = sh.SpecialFolders("Desktop")
        .AssertEqual actual, expected

    .It "should get the path to My Computer"
        actual = ssf.Path( ssf.ssfDRIVES )
        expected = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
        .AssertEqual actual, expected

    .It "should get the path to the Recycle Bin"
        actual = ssf.Path( ssf.ssfBITBUCKET )
        expected = "::{645FF040-5081-101B-9F08-00AA002F954E}"
        .AssertEqual actual, expected

    .It "should all add up"
        actual = _
            ssf.ssfDESKTOP + _
            ssf.ssfPROGRAMS + _
            ssf.ssfCONTROLS + _
            ssf.ssfPRINTERS + _
            ssf.ssfPERSONAL + _
            ssf.ssfFAVORITES + _
            ssf.ssfSTARTUP + _
            ssf.ssfRECENT + _
            ssf.ssfSENDTO + _
            ssf.ssfBITBUCKET + _
            ssf.ssfSTARTMENU + _
            ssf.ssfDESKTOPDIRECTORY + _
            ssf.ssfDRIVES + _
            ssf.ssfNETWORK + _
            ssf.ssfNETHOOD + _
            ssf.ssfFONTS + _
            ssf.ssfTEMPLATES + _
            ssf.ssfCOMMONSTARTMENU + _
            ssf.ssfCOMMONPROGRAMS + _
            ssf.ssfCOMMONSTARTUP + _
            ssf.ssfCOMMONDESKTOPDIR + _
            ssf.ssfAPPDATA + _
            ssf.ssfPRINTHOOD + _
            ssf.ssfLOCALAPPDATA + _
            ssf.ssfALTSTARTUP + _
            ssf.ssfCOMMONALTSTARTUP + _
            ssf.ssfCOMMONFAVORITES + _
            ssf.ssfINTERNETCACHE + _
            ssf.ssfCOOKIES + _
            ssf.ssfHISTORY + _
            ssf.ssfCOMMONAPPDATA + _
            ssf.ssfWINDOWS + _
            ssf.ssfSYSTEM + _
            ssf.ssfPROGRAMFILES + _
            ssf.ssfMYPICTURES + _
            ssf.ssfPROFILE + _
            ssf.ssfSYSTEMx86 + _
            ssf.ssfPROGRAMFILESx86
        expected = &h00 + &h02 + &h03 + &h04 + &h05 + &h06 + _
            &h07 + &h08 + &h09 + &h0a + &h0b + &h10 + &h11 + _
            &h12 + &h13 + &h14 + &h15 + &h16 + &h17 + &h18 + _
            &h19 + &h1a + &h1b + &h1c + &h1d + &h1e + &h1f + _
            &h20 + &h21 + &h22 + &h23 + &h24 + &h25 + &h26 + _
            &h27 + &h28 + &h29 + &h30
    .AssertEqual actual, expected

    .It "should get all the constants"
    .AssertEqual UBound( ssf.AllConstants ) + 1, 38

    .It "should get all the paths"
    .AssertEqual UBound( ssf.AllPaths ) + 1, 38
End With

'Check the command line for /ShowPaths
With WScript.Arguments.Named
    If .Exists( "ShowPaths" ) Then
        ShowPaths
    End If
End With

'For each shell special folder, show the path.
Sub ShowPaths : With WScript.StdOut
    Dim net : Set net = CreateObject("WScript.Network")
    Dim b : b = Chr(8) : b = b & b & b & b 'backspaces
    .WriteLine vbLf & "Shell special folders on this machine for the current user, " & net.UserName
    Dim constant : For Each constant In ssf.AllConstants
        .WriteLine Hex(constant) & vbTab & b & ssf.Path(constant)
    Next
End With : End Sub
