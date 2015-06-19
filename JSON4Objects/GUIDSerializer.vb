Friend Class GUIDSerializer
    Implements ICustomSerializer

    Public Function CanDeserialize(context As DeserializationContext, typeName As String, obj As Object) As Boolean Implements ICustomSerializer.CanDeserialize
        Dim targetType = Type.GetType(typeName)
        Return targetType = GetType(System.Guid)
    End Function

    Public Function CanSerialize(context As SerializationContext, obj As Object, objectsSerializedOnBranch As IDictionary(Of Object, String)) As Boolean Implements ICustomSerializer.CanSerialize
        Return TypeOf obj Is System.Guid
    End Function

    Public Function DeserializeObject(context As DeserializationContext, targetType As System.Type, val As Object,
                                      obj As Object, prop As System.ComponentModel.PropertyDescriptor) As Object Implements ICustomSerializer.DeserializeObject
        If val Is Nothing Then Return Nothing
        Dim guid As System.Guid
        If System.Guid.TryParse(val.ToString().Replace("""", ""), guid) Then
            Return guid
        Else
            Return Nothing
        End If
    End Function

    Public Sub SerializeObject(context As SerializationContext, obj As Object, objectsSerializedOnBranch As IDictionary(Of Object, String)) Implements ICustomSerializer.SerializeObject
        Serializer.SerializeObject(context, obj.ToString(), objectsSerializedOnBranch)
    End Sub
End Class
