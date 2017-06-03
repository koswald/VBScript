
'manual test that the .cs file compiled correctly
'and that the .dll registered correctly

'expected outcome: a synthesized voice says the specified word(s)

With CreateObject("Vox")
    .say "testing"
End With
