Public Class SerializationContext
    Implements ISerializationContext

    Public Property Serializer As Serializer Implements ISerializationContext.Serializer
    Public Property SerializedObjects As Dictionary(Of Type, Dictionary(Of Object, Object))
    Public Property Stream As IWriter
    Public Property Indent As Integer

    Public Sub New(ByVal serializer As Serializer, ByVal stream As IWriter)
        Me.Serializer = serializer
        Me.Stream = stream
        Me.SerializedObjects = New Dictionary(Of Type, Dictionary(Of Object, Object))
        Me.Indent = 0
    End Sub
End Class
