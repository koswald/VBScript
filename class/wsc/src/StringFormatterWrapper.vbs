
'script for StringFormatter.wsc

'expose/wrap the class procedures
Dim sf : Set sf = New StringFormatter
Function Format(arr) : Format = sf.Format(arr) : End Function
Sub SetSurrogate(str) : sf.SetSurrogate str : End Sub
Function Pluralize(i, v) : Pluralize = sf.Pluralize(i, v) : End Function
Sub SetZeroSingular : sf.SetZeroSingular : End Sub
Sub SetZeroPlural : sf.SetZeroPlural : End Sub
