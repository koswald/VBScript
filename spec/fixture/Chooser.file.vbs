
'fixture for Chooser.spec.vbs

'opens a browse for file window and returns user-selected value

'requires the calling script to simulate user action

With CreateObject( "VBScripting.Includer" )
    Execute .Read( "Chooser" )
End With
Dim ch : Set ch = New Chooser
ch.SetBFFileTimeout 5

WScript.StdOut.WriteLine ch.File
