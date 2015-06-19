Public Class DeserializationContext
    Implements ISerializationContext

    Public Property Serializer As Serializer Implements ISerializationContext.Serializer
    Public Property SerializedObjects As Dictionary(Of Object, DeserializedObject)
    Public Property Stream As IO.StreamReader
    Public Property ReferencesToResolve As IList(Of Serializer.ReferenceToResolve)

    Public Sub New(ByVal serializer As Serializer, ByVal stream As IO.StreamReader)
        Me.Serializer = serializer
        Me.Stream = stream
        Me.SerializedObjects = New Dictionary(Of Object, DeserializedObject)
        Me.ReferencesToResolve = New List(Of Serializer.ReferenceToResolve)
    End Sub
End Class

Public Class DeserializedObject
    Public Property Result As Object
    Public Property Origin As Hashtable

    Public Sub New(result As Object, origin As Hashtable)
        Me.Result = result
        Me.Origin = origin
    End Sub
End Class