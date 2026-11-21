' A Python object for the ASF engine: the proxy that asfpy_lo.py gives (an XInvocation), inside a
' Basic object. LibreOffice Basic takes an object that implements XInvocation as an instance of any
' class (TypeOf proxy Is ASF_Map is True), and the engine tells its own objects by TypeOf, so it
' gets this object instead. LOhelper.CallSpecial calls the proxy's members, and wraps the Python
' objects they give.
Public Proxy As Object

' What typeof shows: the class of the Python object
Public Function LOClassName() As String
    LOClassName = Proxy.getImplementationName()
End Function
