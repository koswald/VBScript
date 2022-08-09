'The NameValue class has two properties, Name and Value, which can be used, for example, to describe a startup item in the registry Run key. See the <a href="#startupitems"> StartupItems class</a>.
'
'<pre> With CreateObject( "VBScripting.Includer" )<br />     Execute .Read( "NameValue" )<br /> End With<br /> Set obj = New NameValue.Init( "age", 70 )</pre>
'
Class NameValue

    'Property Name
    'Returns: a variant
    'Parameter: a variant
    Property Let Name(newValue)
        name_ = newValue
    End Property
    Property Get Name
        Name = name_
    End Property
    Private name_

    'Property Value
    'Returns: a variant
    'Parameter: a variant
    Property Let Value(newValue)
        value_ = newValue
    End Property
    Property Get Value
        Value = value_
    End Property
    Private value_

    'Property Init
    'Parameters: name, value
    'Returns an object self reference
    'Remarks: Initializes the object. The Init property returns an object self reference, so an object may be instantiated and initialized in the same statement. See the example. See the <a target="_blank" href="https://github.com/koswald/VBScript/blob/master/class/NameValue.vbs"> code</a>.
    Public Default Property Get Init(n, v)
        Name = n
        Value = v
        Set Init = me
    End Property
    
End Class
