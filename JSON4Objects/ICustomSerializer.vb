
Public Interface ICustomSerializer

    Function CanSerialize(ByVal context As SerializationContext, obj As Object, objectsSerializedOnBranch As IDictionary(Of Object, String)) As Boolean
    Function CanDeserialize(ByVal context As DeserializationContext, typeName As String, obj As Object) As Boolean
    Sub SerializeObject(ByVal context As SerializationContext, ByVal obj As Object, objectsSerializedOnBranch As IDictionary(Of Object, String))
    Function DeserializeObject(ByVal context As DeserializationContext, targetType As Type, ByVal val As Object, obj As Object, prop As ComponentModel.PropertyDescriptor) As Object

End Interface
