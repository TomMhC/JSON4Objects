''' <summary>
''' Set Cancel to True to avoid serialization of this property.
''' Modify Val to change the serialized value.
''' </summary>
Public Class BeforeSerializeEventArgs
    Private _propertyName As String
    ''' <summary>
    ''' The name of the property that will be serialized.
    ''' </summary>
    Public ReadOnly Property PropertyName As String
        Get
            Return _propertyName
        End Get
    End Property

    Private _property As ComponentModel.PropertyDescriptor
    ''' <summary>
    ''' The property that will be serialized. Can be Nothing.
    ''' </summary>
    Public ReadOnly Property [Property] As ComponentModel.PropertyDescriptor
        Get
            Return _property
        End Get
    End Property

    Private _objectType As Type
    ''' <summary>
    ''' Type of the object from which the property is serialized.
    ''' </summary>
    Public ReadOnly Property ObjectType As Type
        Get
            Return _objectType
        End Get
    End Property

    ''' <summary>
    ''' The object from which the property is serialized.
    ''' </summary>
    Public Property Obj As Object

    ''' <summary>
    ''' The value that will be serialized. Modify this to change the actually
    ''' serialized value regardless of the original.
    ''' </summary>
    Public Property Val As Object

    ''' <summary>
    ''' Set Cancel to True to skip this property in serialization.
    ''' </summary>
    Public Property Cancel As Boolean = False

    Public Property Context As SerializationContext

    Public Sub New(propertyName As String, objectType As Type,
                   obj As Object, val As Object,
                   context As SerializationContext)
        _propertyName = propertyName
        _objectType = objectType
        _Obj = obj
        _Val = val
        _Context = context
    End Sub

    Public Sub New([property] As ComponentModel.PropertyDescriptor, objectType As Type,
                   obj As Object, val As Object,
                   context As SerializationContext)
        _property = [property]
        _propertyName = [property].Name
        _objectType = objectType
        _Obj = obj
        _Val = val
        _Context = context
    End Sub
End Class

Public Class HydrateObjectEventArgs

    Public Property Context As DeserializationContext

    Private _obj As Object
    Public ReadOnly Property Obj As Object
        Get
            Return _obj
        End Get
    End Property

    Private _val As Object
    Public ReadOnly Property Val As Object
        Get
            Return _val
        End Get
    End Property

    Private _prop As System.ComponentModel.PropertyDescriptor
    Public ReadOnly Property Prop As System.ComponentModel.PropertyDescriptor
        Get
            Return _prop
        End Get
    End Property

    Public Property Handled As Boolean = False

    Public Sub New(obj As Object, val As Object, prop As System.ComponentModel.PropertyDescriptor,
                   context As DeserializationContext)
        _obj = obj
        _val = val
        _prop = prop
        _Context = context
    End Sub

End Class


Public Class TransformObjectEventArgs

    Private _obj As Object
    Public ReadOnly Property Obj As Object
        Get
            Return _obj
        End Get
    End Property

    Public Property Val As Object

    Public Property Handled As Boolean = False

    Public Sub New(obj As Object, val As Object)
        _obj = obj
        Me.Val = val
    End Sub

End Class


Public Class NewInstanceEventArgs

    Public Property NewObj As Object

    Public Property Handled As Boolean = False

    Public Property Context As DeserializationContext

    Public Property TargetType As Type

    Public Property SourceHashTable As Hashtable

    Public Sub New(context As DeserializationContext, targetType As System.Type, sourceHashTable As Hashtable)
        Me.Context = context
        Me.TargetType = targetType
        Me.SourceHashTable = sourceHashTable
    End Sub

End Class

Public Class ObjectDeserializedEventArgs

    Public Property Context As DeserializationContext
    Public Property Result As Hashtable

    Public Sub New(context As DeserializationContext, result As Hashtable)
        Me.Context = context
        Me.Result = result
    End Sub

End Class