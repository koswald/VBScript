
'fixture for Chooser.spec.vbs

'opens a browse for file window and returns user-selected value

'requires the calling script to simulate user action

With CreateObject("includer")
    Execute(.read("Chooser"))
End With
Dim ch : Set ch = New Chooser

WScript.StdOut.WriteLine ch.File
