
'fixture for Chooser.spec.vbs

'opens a browse for folder window and returns user-selected value

'requires the calling script to simulate user action

'specifically,
'when the user cancels out of the dialog, 
'then the following statement is True
'TypeName(obj) = "Nothing"

With CreateObject("includer")
    Execute .read("Chooser")
End With
Dim ch : Set ch = New Chooser

ch.SetRootPath "%tmp%"
Dim obj : Set obj = ch.FolderObject

WScript.StdOut.WriteLine TypeName(obj)

Set obj = Nothing
