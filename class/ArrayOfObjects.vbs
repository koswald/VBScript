'The default property of the ArrayOfObjects class, Items, acts like a rudimentary C# ArrayList.
'
'Example
'<pre> Option Explicit<br /> Dim aoo 'ArrayOfObjects object<br /> Dim incl 'VBScripting.Includer object<br /> Initialize<br /> Add "tree", "pear"<br /> Add "tree", "walnut"<br /> ShowAll<br /> ShowAll2<br /> Sub Initialize<br />     Set incl = CreateObject( "VBScripting.Includer" )<br />     Execute incl.Read( "ArrayOfObjects" )<br />     Set aoo = New ArrayOfObjects<br /> End Sub<br /> Sub Add( noun, example )<br />     Execute incl.Read( "NameValue" )<br />     aoo.Add New NameValue.Init( noun, example )<br /> End Sub<br /> Sub ShowAll<br />     Dim obj, s<br />     For Each obj In aoo() 'or aoo.Items or aoo.Items()<br />         s = s & obj.Name & vbTab & obj.Value & vbLf<br />     Next<br />     MsgBox s,, "ShowAll"<br /> End Sub<br /> Sub ShowAll2<br />     Dim i, s<br />     For i = 0 To UBound(aoo) 'or aoo() or aoo.Items or aoo.Items()<br />         s = s & aoo()(i).Name & vbTab & aoo()(i).Value & vbLf<br />     Next<br />     MsgBox s,, "ShowAll2"<br /> End Sub</pre>
'
Class ArrayOfObjects

    Private items_() 'dynamic array
    Private count_  'integer

    Sub Class_Initialize
        Count = 0
        ReDim items_(-1)
    End Sub

    'Property Items
    'Returns an array of objects
    'Remarks: Returns an array of the objects that were added using the Add method. This is default property, so the name (Items) may not need to be specified. However, it may be necessary to add empty parens to the object name: See the example.
    Public Default Property Get Items
        Items = items_
    End Property

    'Method Add
    'Parameter: an object
    'Remarks: Expands the Items array and adds the specified object to it.
    Sub Add(item)
        ReDim Preserve items_(Count)
        Set items_(Count) = item
        Count = Count + 1
    End Sub

    'Property Count
    'Returns an integer
    'Remarks: Returns the number of items in the Items array.
    Public Property Get Count
        Count = count_
    End Property
    Private Property Let Count(newCount)
        count_ = newCount
    End Property

End Class
