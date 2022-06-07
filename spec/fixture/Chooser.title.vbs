
'fixture for Chooser.spec.vbs

'opens a browse for folder window and returns user-selected value

'requires the calling script to simulate user action

With CreateObject( "VBScripting.Includer" )
    Execute .Read( "Chooser" )
End With
Dim ch : Set ch = New Chooser

ch.SetRootPath "%tmp%"
WScript.StdOut.WriteLine ch.FolderTitle
